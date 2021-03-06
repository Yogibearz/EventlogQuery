#Requires -version 2.0
#
# powershell -command "& {./EventlogQuery.1.9.9.9.ps1}"
# C:\WINDOWS\Microsoft.NET\Framework\v1.1.4322\gacutil.exe  "%HOME%\My Documents\WindowsPowerShell\log4net.dll"
# %UserProfile%\My Documents\WindowsPowerShell\Microsoft.PowerShell_profile.ps1
# Set-ExecutionPolicy RemoteSigned

set-psdebug -strict

$ErrorActionPreference = 'Stop'
$version = "1.9.9.9"
$configF = "EventlogQueryConfig.xml"
$log = $null
$pwd = $(Get-Location)
$begin = get-date

if (-not(Test-Path $configF)) { write-host "$configF missing"; exit -999}

[xml]$R = Get-Content $configF

$debug = $R.config.debug

#if ($debug){ Set-Location -Path f:\temp\don }

if ($R.config.version -ne $version) { write-host "wrong version of configuration.xml"; exit -999}

if ((Test-Path "$env:userprofile\My Documents\WindowsPowerShell\log4net.dll") -and (Test-Path "$pwd\$($R.config.log.logconfig)")) {
   [System.Reflection.Assembly]::LoadFrom("$env:userprofile\My Documents\WindowsPowerShell\log4net.dll") | out-null
   $log4netconfigfile = new-Object System.Io.FileInfo("$pwd\$($R.config.log.logconfig)")
   [log4net.LogManager]::ResetConfiguration()
   [log4net.GlobalContext]::Properties["PWD"] = $pwd
   [log4net.Config.XmlConfigurator]::ConfigureAndWatch($log4netconfigfile)
   $log = [log4net.LogManager]::GetLogger("root")
   $log.info("== Job Start $version ==")
} else {
	 Write-host "log4net.dll or $($R.config.log.logconfig) missing"
	 exit -100
}

#######################################################################################
#  Main
#######################################################################################

$messages = @()
$count = 0

$Timestamps = @{}
$TimestampsFile = $R.config.log.timestamp
$dateformat = $R.config.log.attachpostfix
$outputT = (get-date).toString($dateformat)
$oldcuttime = $null
$logdatetime = $R.config.log.datetime
$htmlfileage = $R.config.oldfiles.htmlfileage
$timestampfileage = $R.config.oldfiles.timestampfileage


###
# Check OS
###

#$los = Get-WmiObject -Query "Select Caption from Win32_OperatingSystem" -EnableAllPrivileges

###
# Load back last scan time for each machine
###


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

[int[]]$Events = $($R.config.filter.eventid).split(",")

###
# Check each machine
###

$computers = @()
$strF = @()
$strNT = @()

if ($debug -ne '0'){
	 [int[]]$Events += $($R.config.filter.debugeventid).split(",")
	 $computers = ('U-S19', 'U-S20','U-024813')
	 #For ($i=91; $i -le 95; $i++) { $computers += ('U-VW' + (7000 + $i)) }
	 #For ($i=375; $i -le 375; $i++) { $computers += ('U-VW' + (7000 + $i)) }
} else {
	 For ($i=1; $i -le 500; $i++) { $computers += ('U-VW' + (7000 + $i)) }
	 For ($i=1; $i -le 20; $i++) { $computers += ('U-VX' + $i.ToString("0000")) }
}

$Events = $Events | Sort-Object

Foreach ($Machine in $computers) {
   $mbegin = get-date
   $ErrorActionPreference = "SilentlyContinue"
   $resolve = ""
   $resolve = [System.Net.Dns]::GetHostAddresses($Machine)
   
   if (ping($Machine) -and $resolve -ne $Null) {

      #if windows 7
      #$ServiceStatus = (Get-WmiObject -computername $Machine -class win32_service -Filter "Name='RemoteRegistry'").Status
      #if ($ServiceStatus -eq "Stopped") {
      #   (Get-WmiObject -computername $Machine -class win32_service -Filter "Name='RemoteRegistry'").StartService()
      #}

      $ErrorActionPreference = "SilentlyContinue"
      #if ($resolve -eq $Null) { 
      #	 $resolve = ""
      #	 $reverseResolve = ""
      #} else {
      	 $reverseResolve = ""
      	 $reverseResolve = [System.Net.Dns]::GetHostEntry($resolve).Hostname
      #}	 	
       
      $log.info(("{0} : {1} : {2}" -f $Machine,$resolve,$reverseResolve))
      #if ($resolve -match "10*") { Continue }

      $cuttime = [datetime]::ParseExact("20000101-000000", $dateformat, $null)
      #$cuttime = [datetime]::ParseExact("20120815-000000", $dateformat, $null)
      if ($Timestamps.ContainsKey($Machine)) {
         $oldcuttime = $Timestamps["$Machine"]
         #$log.debug("$Machine " + $Timestamps["$Machine"] + " $dateformat")
         $cuttime = [datetime]::ParseExact($Timestamps["$Machine"], $dateformat, $null)
         $log.Debug("$Machine old cut time [" + $Timestamps["$Machine"] + "]")
      } else {
      	 $oldcuttime = $null
      }

      $Timestamps["$Machine"] = (get-date).toString($dateformat)
      $log.Debug("$Machine new cut time [" + $Timestamps["$Machine"] + "]")


      $CatchFlag = 0
      $result = @()

      $strCuttime = [System.Management.ManagementDateTimeConverter]::ToDMTFDateTime($cuttime)

      $log.debug("Filter : TimeWritten > $strCuttime")

      $filter  = "logfile='$($R.config.filter.eventlogfile)' AND TimeWritten > '$strCuttime' "
      $filter += "AND (EventCode=" + [string]::join(" OR EventCode=", $Events) + ")"

      Try {
         #$log.debug($filter)
         #Measure-Command {
         $result = Get-WmiObject -EA SilentlyContinue -Computer $Machine -Class "Win32_NTLogEvent" -Filter $filter
         #}
         if (! $?) {
         	  $CatchFlag = -1
         	  $log.debug("CatchFlag [$CatchFlag]")
         	  throw $error[0].Exception
         }
      }
      #Catch [System.UnauthorizedAccessException]{
      Catch {
         $ErrorMessage = $_.Exception.Message
         $FailedItem = $_.Exception.ItemName
         $ErrorMessage = $ErrorMessage -replace "`t|`n|`r",""
         $log.error("*Fail to get $Machine eventlog [$ErrorMessage]")
         $CatchFlag = 1
      }
      Finally {
         if ($Error[0].Exception -match "HRESULT: 0x800706BA") {
         	   $log.error("Fail to get $Machine eventlog [WMI COM (RPC) not available]")
         	   $CatchFlag = -1
         }
         $log.debug("[F] CatchFlag [$CatchFlag] $Machine count [" + $result.count + "]")
         if ($CatchFlag -eq 0 -or ([bool]$result.count)) {
            $log.debug("$Machine count [" + $result.count + "]")
            $count += $result.count
            $messages += $result
            $log.debug("Total count [" + $messages.count + "]")
         }
      }
      $ErrorActionPreference = 'Stop'
      if ($CatchFlag -ne 0 -and (-not [bool]$result.count)) {
      	 if ($oldcuttime -eq $null) {
      	 	  $Timestamps.remove($Machine)
      	 	  $log.debug("$Machine reset cuttime as none")
      	 }
      } else {
      	 if ([bool]$result.count) {
      	    # get newest record time
      	    $sortRecords = ($result | Sort-Object -descending TimeGenerated)[0].TimeGenerated
      	    #$Timestamps["$Machine"] = (([system.management.managementdatetimeconverter]::todatetime($sortRecords)) -format $dateformat)
      	    $Timestamps["$Machine"] = [system.management.managementdatetimeconverter]::todatetime($sortRecords).toString($dateformat)
      	    $log.debug("$Machine set cuttime {0}" -f $Timestamps["$Machine"])
      	 }   
      	 $strF += ("{0} {1}" -f $Machine,$filter)
      	 $strNT += ("{0} [{1}] [{2}]" -f $Machine,$oldcuttime,$Timestamps["$Machine"])
         $mend = get-date
         $mts = $mend - $mbegin
         $log.info("$Machine Process time [" + ('{0:00}:{1:00}:{2:00}' -f $mts.Hours,$mts.Minutes,$mts.Seconds) + "]")
      }
   } else {
   	  if ($resolve -eq $Null) {
   	  	 $log.error("$Machine has no A record")
   	  } else {
   	  	 $log.info("$Machine not alive")
   	  }
      if ($oldcuttime -eq $null) {
      	  $Timestamps.remove($Machine)
      	  $log.debug("$Machine reset cuttime as none")
      } else {
         $Timestamps["$Machine"] = $oldcuttime
         $log.debug("$Machine reset cuttime $oldcuttime")
      }
   }

   # for Windows 7
   #if ($ServiceStatus -eq "Stopped") {
   #	  (Get-WmiObject -computername $Machine -class win32_service -Filter "Name='RemoteRegistry'").StopService()
   #}
}
$log.info("Matched eventlogs [$count]")

#$groupbyA = $messages | Group-Object ComputerName,EventCode
$groupby = $messages | Group-Object ComputerName,EventCode | Sort-Object Name | `
           ConvertTo-HTML -Fragment -property @{LABEL='Name'; EXPRESSION={$_.Name.ToUpper().Replace('.U-.COM','').split(", ")[0]}}, `
           @{LABEL='Event ID';EXPRESSION={$_.Name.split(", ")[2]}}, Count | `
           foreach {$_.replace("</table>","</table>`n<p><br /></p>").replace("<table>","<table class='group'>")}

###
# Save last scan time for each machine and total scan result
###

$oHead = @"
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>Eventlog Scan Result : $outputT</title>
<style type="text/css">
BODY{ background-color: #FFFFFF;
      font-family: Arial,Tahoma,sans-serif;
      font-weight: 400;
      font-style: normal;
      font-size: 12pt; }
TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse; padding:25px; width: 100%;font-size: 10pt;}
TH{border-width: 1px;padding: 2px;border-style: solid;border-color: black;background-color:#F4D8B1}
TD{border-width: 1px;padding: 2px;border-style: solid;border-color: black;background-color:#FFFFBE;text-align: center;}
.info {font-size: 16px; line-height: 24px; }
.process {text-shadow: 0px 2px 3px #555;}
.EventR {color: Navy; font-weight:bold;}
.Ptime {color: CornflowerBlue; font-weight:bold;}
.machine {color: DarkOrange; font-weight:bold; }
.gen {font-size: 8pt;}
.id {color: red;  font-weight:bold; background-color:greenyellow;}
TABLE TR:hover TD { background:#BAFECB !important; }
TABLE.group {vertical-align: middle;width: auto;}
TABLE.group TD,TH {vertical-align: middle;text-align: center; padding:2px 10px 2px 10px;}
</style>
"@


$end = get-date
$ts = $end - $begin

$oPost  = ("<p><span class='info'><br />Event ID <span class='id'>{0}</span> Scanned`n" -f ($Events -join ","))
$oPost += ("<br /><span class='process'><span class='machine'>{4}</span> queries executed,`n Total Process Time <span class='Ptime'>{0:00}:{1:00}:{2:00}</span>,`n <span class='EventR'>{3} Events Returned</span></span><br />`n" -f $ts.Hours,$ts.Minutes,$ts.Seconds,$count,$computers.count)
$oPost += ("<br /><span class='gen'>Generated at {0}</span><br /></span></p>`n" -f (get-date).ToString($dateformat + ".ffff"))
$oPost += ("<!-- {0} -->`n" -f $version)
$oPost += ($strF | Foreach-Object -process {"<!-- {0} -->`n" -f $_})
$oPost += ($strNT | Foreach-Object -process {"<!-- {0} -->`n" -f $_})

$attach = "$pwd\{0}{1}.html" -f $R.config.log.attachprefix, $outputT
#$log.debug(("$attach {0} {1}" -f $R.config.log.attachprefix, $outputT))

#if ($messages.count -lt 10) { $groupby = "<br />" }

if ($count -gt 0) {
   #$messages | Sort-Object ComputerName, TimeGenerated | ConvertTo-HTML -Title "$outputT" -head $oHead -Body $oPost `
   #     -property ComputerName, @{LABEL="TimeGenerated"; EXPRESSION = {$_.convertToDateTime($_.TimeGenerated)}}, `
   #     EventCode, EventType, SourceName, Message, RecordNumber > "$pwd\Filter-$outputT.html"
   $messages | Sort-Object ComputerName, TimeGenerated | ConvertTo-HTML -Title "$outputT" -head $oHead -Body $oPost `
        -PreContent $groupby `
        -property @{LABEL="ComputerName"; EXPRESSION = {$_.ComputerName.ToUpper().Replace('.U-.COM','')}}, `
        @{LABEL="TimeGenerated"; EXPRESSION = {"{0:$logdatetime}" -f $_.convertToDateTime($_.TimeGenerated)}}, `
        EventCode, EventType, SourceName, Message | Set-Content -Encoding UTF8 "$attach"
        #foreach-object {$_.replace("<table>","<table id='maintable'>")} | `

   if ($debug -and (Test-Path "$attach")) { Invoke-Item "$attach" }
   $log.debug("Matched eventlogs HTML saved")
}

#Get-Item "$pwd\$TimestampsFile" | Rename-Item $_ "$($_.DirectoryName)\$($_.BaseName).$($_.LastWriteTime.toString($dateformat))$($_.extension)"
Get-Item "$pwd\$TimestampsFile" | foreach-object {$newName = "$($_.DirectoryName)\$($_.BaseName).$($_.LastWriteTime.toString($dateformat))$($_.extension)" ; Rename-Item $_.Fullname -newname $newName }

$Timestamps.GetEnumerator() | Select Key,Value,@{Name="Type";Expression={$_.value.gettype().name}} | Sort-Object Name | Export-CSV -path "$pwd\$TimestampsFile" -force
if ($count -gt 0) {
   $log.debug("Cut Time CSV $pwd\$TimestampsFile Saved ({0:$dateformat})" -f (Get-Item "$pwd\$TimestampsFile").LastWriteTime)
}

###
# Mail out result
###

$emailFrom = "VDI_Scan@u-um.com"
$emailTo = $($R.config.mail.to).split(",")
$subject = "Suspicious Events ($count)"
$smtpServer = $R.config.mail.smtpserver
$CatchFlag = 0

if ($count -gt 0) {
   try {
      Send-MailMessage -To $emailTo -Subject $subject -Body " " `
      	-SmtpServer $smtpServer -From $emailFrom -Attachment "$attach" `
      	-DeliveryNotificationOption OnFailure -ErrorAction Continue
   } catch {
      $ErrorMessage = $_.Exception.Message
      $FailedItem = $_.Exception.ItemName
      $ErrorMessage = $ErrorMessage -replace "`t|`n|`r",""
      $log.error("Fail to send mail [$ErrorMessage]")
      $CatchFlag = 1
   }
   if (! $CatchFlag) {$log.info("Mail sent")}
} else {
   try {
	    Send-MailMessage -To $emailTo -Subject $subject -Body " " `
      		-SmtpServer $smtpServer -From $emailFrom `
      		-DeliveryNotificationOption OnFailure -ErrorAction Continue
   } catch {
      $ErrorMessage = $_.Exception.Message
      $FailedItem = $_.Exception.ItemName
      $ErrorMessage = $ErrorMessage -replace "`t|`n|`r",""
      $log.error("Fail to get $Machine eventlog [$ErrorMessage]")
      $CatchFlag = 1
   }
   if (! $CatchFlag) {$log.info("*** No matched eventlog entry")}
}

###
# Remove old scan result files
###

$LastWrite = $([datetime]::ParseExact($outputT, $dateformat, $null)).AddDays(-$htmlfileage)
$OldFiles = Get-Childitem $pwd -Include "$($R.config.log.attachprefix)*.html" | Where {$_.LastWriteTime -le "$LastWrite"}

foreach ($XFile in $OldFiles) {
   if ($XFile -ne $NULL) {
       $log.info("Delete File $XFile")
       #Remove-Item $XFile.FullName #| out-null
   } else {
       Write-Host "No more files to delete!" -foregroundcolor "Green"
   }
}

###
# Remove old timestamp files
###

$LastWrite = $([datetime]::ParseExact($outputT, $dateformat, $null)).AddDays(-$timestampfileage)
$OldFiles = Get-Childitem "$pwd\*.*" -Include "$($TimestampsFile.split(".")[0])*.$($TimestampsFile.split(".")[1])" -Exclude $TimestampsFile | Where {$_.LastWriteTime -le "$LastWrite"}

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
