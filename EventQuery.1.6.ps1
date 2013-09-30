#
# V 1.6
#
# powershell -command "& './EventlogQuery.1.6.ps1'"
# C:\WINDOWS\Microsoft.NET\Framework\v1.1.4322\gacutil.exe  "%HOME%\My Documents\WindowsPowerShell\log4net.dll"
# %UserProfile%\My Documents\WindowsPowerShell\Microsoft.PowerShell_profile.ps1

set-psdebug -strict
$ErrorActionPreference = 'Stop'
$log = $null

if ((Test-Path "$pwd\log4net.dll") -and (Test-Path "$pwd\log4net.config")) {
   #$log = New-Logger -Configuration "$pwd\log4net.config" -Dll "$pwd\log4net.dll" -Verbose
   #$log.DebugFormat("Logger configuration file is : '{0}'", (Resolve-Path "$pwd\log4net.config"))
   #$log.InfoFormat("test test {0}", "log4test")
   [System.Reflection.Assembly]::LoadFrom("$pwd\log4net.dll") | out-null
   $log4netconfigfile = new-Object System.Io.FileInfo("$pwd\log4net.config")
   [log4net.LogManager]::ResetConfiguration()
   [log4net.Config.XmlConfigurator]::ConfigureAndWatch($log4netconfigfile)
   $log = [log4net.LogManager]::GetLogger("root")
   $log.info("== Job Start ==")
} else {
	 Write-host "log4net.dll or log4net.config missing"
	 exit -100
}

#######################################################################################
#  Main
#######################################################################################

$pwd = "."
$messages = @()
$count = 0

$computers = ('UMCS18', 'UMCS20', 'UMCS19')
$Timestamps = @{}
$TimestampsFile = 'Timestamps.csv'
$dateformat = "yyyyMMdd-HHmmss"
$outputT = get-date -format $dateformat

$a = "<style>"
$a = $a + "BODY{background-color:peachpuff;}"
$a = $a + "TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse; padding:25px;}"
$a = $a + "TH{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color:thistle}"
$a = $a + "TD{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color:palegoldenrod}"
$a = $a + "</style>"

###
# Load back last scan time for each machine
###

#$hash = @{
#  Server           = ""
#  Datetime         = $(get-date -format yyyyMMdd-HHmmss)
#}
#$Timestamps = New-Object PSObject -Property $hash
$Timestamps = @{}

if (Test-Path "$pwd\$TimestampsFile") {
	#$Timestamps = Import-CSVtoHash "$pwd\$TimestampsFile"
	$log.Debug("Cut Time File exists")
	$ErrorActionPreference = 'Continue'

	try {
    	Import-Csv -Path "$pwd\$TimestampsFile" | ForEach-Object {`
          $Timestamps.Add($_.Key, $($_.Value))
          #$Timestamps["$_.Key"] = $([datetime]::ParseExact($($_.Value), $dateformat, $null))
          }
  }
  catch {
      $ErrorMessage = $_.Exception.Message
      $FailedItem = $_.Exception.ItemName
      $ErrorMessage = $ErrorMessage -replace "`t|`n|`r",""
      $log.error("Fail to load Cut Time File : $ErrorMessage")
      exit -200
  }

  $ErrorActionPreference = 'Stop'
  $log.Info("Cut Time Loaded")
}

###
# Check each machine
###

#$wmi = Get-WmiObject win32_service -filter "name = 'spooler'"

#Foreach ($Machine in $computers) {
For ($i=1; $i -le 500; $i++) {
   $Machine = 'UMCVW' + (7000 + $i)

   $cuttime = [datetime]::ParseExact("19000101-000000", $dateformat, $null)
   #$cuttime = [datetime]::ParseExact("20120815-000000", $dateformat, $null)
   if ($Timestamps.ContainsKey($Machine)) {
      $cuttime = [datetime]::ParseExact($Timestamps["$Machine"], $dateformat, $null)
      $log.Debug("$Machine old cut time [" + $Timestamps["$Machine"] + "]")
   }
   
   $Timestamps["$Machine"] = $(get-date -format $dateformat)
   $log.Debug("$Machine new cut time [" + $Timestamps["$Machine"] + "]")

   $ErrorActionPreference = 'SilentContinue'

   $CatchFlag = 0
   $log.debug($cuttime.ToString("yyyyMMddHHmmss.000000+480"))
   $strCuttime = [System.Management.ManagementDateTimeConverter]::ToDMTFDateTime($cuttime)
   try {
      #$message = Get-EventLog -Logname System -computer $Machine -After $cuttime | `
      #           Where-Object {$_.EventId -eq 7 -or $_.EventId -eq 14 -or $_.EventId -eq 41 `
      #           							 -or $_.EventId -eq 23 -or $_.EventId -eq 110 `
      #           							 -or $_.EventId -eq 1117 -or $_.EventId -eq 6072 -or $_.EventId -eq 6013}
      $message = Get-WmiObject -Computer $Machine -Class Win32_NTLogEvent `
                    -Filter "(logfile='System') AND (TimeGenerated > '$strCuttime') AND (EventCode=7 OR EventCode=14 OR EventCode=41 OR EventCode=1117 OR EventCode=6072 OR EventCode=6013)"
   }
   Catch {
      $ErrorMessage = $_.Exception.Message
      $FailedItem = $_.Exception.ItemName
      $ErrorMessage = $ErrorMessage -replace "`t|`n|`r",""
      $log.error("Fail to get $Machine eventlog [$ErrorMessage]")
      $CatchFlag = 1
   }
   Finally {
      if ($CatchFlag = 0) {
         $log.debug("$Machine count [$message.count]")
         $count += $message.count
         $messages += $message
      }
      $CatchFlag = 0
   }

   $ErrorActionPreference = 'Stop'

}

###
# Save last scan time for each machine and total scan result
###

#($messages | ConvertTo-XML -NoTypeInformation).Save("$pwd\Filter-$outputT.xml")
$messages | ConvertTo-HTML -property ComputerName, `
     @{LABEL="TimeGenerated"; EXPRESSION = {$_.convertToDateTime($_.TimeGenerated)}}, `
     EventCode, EventType, SourceName, Message, RecordNumber -head $a > "$pwd\Filter-$outputT.html"
$log.debug("Matched eventlogs HTML saved")
$Timestamps.GetEnumerator() | Select Key,Value,@{Name="Type";Expression={$_.value.gettype().name}} | Export-CSV "$pwd\$TimestampsFile"
$log.debug("Cut Time CSV Saved")

###
# Mail out result
###
if ($count -gt 0) {
   $emailFrom = "donal_sun@umc.com"
   $emailTo = "donal_sun@umc.com"
   $subject = "Suspicious Events ($count)"
   #$messages = Get-ChildItem UMC*.html -name $workpath | foreach ($a in $vara) {Get-Content $files -Delimiter ([char]0)}
   $smtpServer = "202.14.12.4"

   Send-MailMessage -To $emailTo -Subject $subject -Body " " `
   								 -SmtpServer $smtpServer -From $emailFrom -Attachment "$pwd\Filter-$outputT.html" `
   								 -DeliveryNotificationOption OnFailure
   $log.info("Mail sent")
} else {
	 $log.info("*** No matched eventlog")
}

###
# Remove old scan files
###

$Days = 7
$LastWrite = $([datetime]::ParseExact($outputT, $dateformat, $null)).AddDays(-$Days)
$OldFiles = Get-Childitem $pwd -Include "Filter-*.xml" | Where {$_.LastWriteTime -le "$LastWrite"}

foreach ($XFile in $OldFiles) {
   if ($XFile -ne $NULL) {
       $log.info("Delete File $XFile")
       #Remove-Item $XFile.FullName #| out-null
   } else {
       Write-Host "No more files to delete!" -foregroundcolor "Green"
   }
}

$log.info("== Job End ==")
