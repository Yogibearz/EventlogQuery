#Requires -version 2.0
#
# V 1.9.6
#
# powershell -command "& './EventlogQuery.1.9.6.ps1'"
# C:\WINDOWS\Microsoft.NET\Framework\v1.1.4322\gacutil.exe  "%HOME%\My Documents\WindowsPowerShell\log4net.dll"
# %UserProfile%\My Documents\WindowsPowerShell\Microsoft.PowerShell_profile.ps1
# Set-ExecutionPolicy RemoteSigned

set-psdebug -strict

if ("UMCS20UMCS19" -like "$env:computername*" ){ Set-Location -Path f:\temp\don }
	
$ErrorActionPreference = 'Stop'
$version = "1.9.5"
$log = $null
$pwd = "."
$begin = get-date

if ((Test-Path "$env:userprofile\My Documents\WindowsPowerShell\log4net.dll") -and (Test-Path "$pwd\log4net.config")) {
   [System.Reflection.Assembly]::LoadFrom("$env:userprofile\My Documents\WindowsPowerShell\log4net.dll") | out-null
   $log4netconfigfile = new-Object System.Io.FileInfo("$pwd\log4net.config")
   [log4net.LogManager]::ResetConfiguration()
   [log4net.Config.XmlConfigurator]::ConfigureAndWatch($log4netconfigfile)
   $log = [log4net.LogManager]::GetLogger("root")
   $log.info("== Job Start $version ==")
} else {
	 Write-host "log4net.dll or log4net.config missing"
	 exit -100
}

#######################################################################################
#  Main
#######################################################################################

$messages = @()
$count = 0

$computers = ('UMCS19', 'UMCS20')
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
# Check OS
###

$los = Get-WmiObject -Query "Select Caption from Win32_OperatingSystem" -EnableAllPrivileges

###
# Load back last scan time for each machine
###

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

$filter = "(logfile='System') AND (TimeGenerated > '$strCuttime') AND (EventCode=7 OR EventCode=14 OR EventCode=41 OR EventCode=1117 OR EventCode=6072)"
#Foreach ($Machine in $computers) {
For ($i=1; $i -le 500; $i++) {
   $Machine = 'UMCVW' + (7000 + $i)

   $mbegin = get-date
   if (ping($Machine)) { 
      
      #if windows 7
      #$ServiceStatus = (Get-WmiObject -computername $Machine -class win32_service -Filter "Name='RemoteRegistry'").Status
      #if ($ServiceStatus -eq "Stopped") {
      #   (Get-WmiObject -computername $Machine -class win32_service -Filter "Name='RemoteRegistry'").StartService()
      #}
      
      $cuttime = [datetime]::ParseExact("19000101-000000", $dateformat, $null)
      #$cuttime = [datetime]::ParseExact("20120815-000000", $dateformat, $null)
      if ($Timestamps.ContainsKey($Machine)) {
         $cuttime = [datetime]::ParseExact($Timestamps["$Machine"], $dateformat, $null)
         $log.Debug("$Machine old cut time [" + $Timestamps["$Machine"] + "]")
      }
      
      $Timestamps["$Machine"] = $(get-date -format $dateformat)
      $log.Debug("$Machine new cut time [" + $Timestamps["$Machine"] + "]")
      
      $ErrorActionPreference = "SilentlyContinue"
      
      $CatchFlag = 0
      $result = @()
      
      $strCuttime = [System.Management.ManagementDateTimeConverter]::ToDMTFDateTime($cuttime)
      
      $log.debug("TimeGenerated > $strCuttime")
       
      try {
         #$result = Get-WmiObject -Computer $Machine -Class "Win32_NTLogEvent" -Filter "(logfile='System') AND (TimeGenerated > '$strCuttime') AND (EventCode=7 OR EventCode=14 OR EventCode=41 OR EventCode=1117 OR EventCode=6072 OR EventCode=6013 OR EventCode=1085)"
         #$result = Get-WmiObject -Computer $Machine -Query "select * from Win32_NTLogEvent where (logfile='System') AND (TimeGenerated > '$strCuttime')" `
         #            | where-object {$_.EventCode -eq 7 -OR $_.EventCode -eq 14 -OR $_.EventCode -eq 41 -OR $_.EventCode -eq 1117 `
         #           -OR $_.EventCode -eq 6072 -OR $_.EventCode -eq 6013 -OR $_.EventCode -eq 1085}
         $rawlog = Get-WmiObject -EA SilentlyContinue -Computer $Machine -Class "Win32_NTLogEvent" -Filter "logfile='System' AND (EventCode=7 OR EventCode=14 OR EventCode=41 OR EventCode=1117 OR EventCode=6072 OR EventCode=6013)"
         #$result = $rawlog | where-object {$_TimeGenerated -gt '$strCuttime' -and ($_.EventCode -eq 7 -OR $_.EventCode -eq 14 -OR $_.EventCode -eq 1117 `
         #           -OR $_.EventCode -eq 6072 -OR $_.EventCode -eq 6013 -OR $_.EventCode -eq 1085)}
         if (! $?) {
         	  $CatchFlag = -1
         	  $log.debug("CatchFlag [$CatchFlag]")
         } else {
         	  $result = $rawlog | where-object {$_.TimeGenerated -gt '$strCuttime'}
         }
      }
      Catch [System.UnauthorizedAccessException]{
         $ErrorMessage = $_.Exception.Message
         $FailedItem = $_.Exception.ItemName
         $ErrorMessage = $ErrorMessage -replace "`t|`n|`r",""
         $log.error("Fail to get $Machine eventlog [$ErrorMessage]")
         $CatchFlag = 1
      }
      Finally {
         if ($Error[0].Exception -match "HRESULT: 0x800706BA") {
         	   $log.error("Fail to get $Machine eventlog [WMI COM (RPC) not available]")
         	   $CatchFlag = -1
         }
         $log.debug("[F] CatchFlag [$CatchFlag] $Machine count [" + $result.count + "]")
         if ($CatchFlag -eq 0) {  
            $log.debug("$Machine count [" + $result.count + "]")
            $count += $result.count
            $messages += $result
            $log.debug("Total count [" + $messages.count + "]")
         }         
         $CatchFlag = 0
      }
      $ErrorActionPreference = 'Stop'
      
   } else {
   	  $log.info("$Machine not alive")
   }
   
   # for Windows 7
   #if ($ServiceStatus -eq "Stopped") {
   #	  (Get-WmiObject -computername $Machine -class win32_service -Filter "Name='RemoteRegistry'").StopService()
   #}
   $mend = get-date
   $mts = $mend - $mbegin
   $log.info("$Machine Process time [" + ('{0:00}:{1:00}:{2:00}' -f $mts.Hours,$mts.Minutes,$mts.Seconds) + "]")

}
$log.info("Matched eventlogs [$count]")

###
# Save last scan time for each machine and total scan result
###

if ($count -gt 0) {
   #($messages | ConvertTo-XML -NoTypeInformation).Save("$pwd\Filter-$outputT.xml")
   $messages | ConvertTo-HTML -property ComputerName, `
        @{LABEL="TimeGenerated"; EXPRESSION = {$_.convertToDateTime($_.TimeGenerated)}}, `
        EventCode, EventType, SourceName, Message, RecordNumber -head $a > "$pwd\Filter-$outputT.html"
   $log.debug("Matched eventlogs HTML saved")
}
$Timestamps.GetEnumerator() | Select Key,Value,@{Name="Type";Expression={$_.value.gettype().name}} | Export-CSV "$pwd\$TimestampsFile"
$log.debug("Cut Time CSV Saved")

###
# Mail out result
###
if ($count -gt 0) {
   $emailFrom = "VDI_Scan@umc.com"
   $emailTo = "arvin_ch_lin@umc.com,donal_sun@umc.com"
   $subject = "Suspicious Events ($count)"
   #$messages = Get-ChildItem UMC*.html -name $workpath | foreach ($a in $vara) {Get-Content $files -Delimiter ([char]0)}
   $smtpServer = "F12AG01"

   Send-MailMessage -To $emailTo -Subject $subject -Body " " `
   								 -SmtpServer $smtpServer -From $emailFrom -Attachment "$pwd\Filter-$outputT.html" `
   								 -DeliveryNotificationOption OnFailure -ErrorAction Continue
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

$end = get-date
$ts = $end - $begin
$log.info("Process time [" + ('{0:00}:{1:00}:{2:00}' -f $ts.Hours,$ts.Minutes,$ts.Seconds) + "]")
$log.info("== Job End ==")
