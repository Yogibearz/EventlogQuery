#
# V 1.2
#

set-psdebug -strict

function Join-Object {
    Param(
       [Parameter(Position=0)]
       $First
    ,
       [Parameter(Position=1,ValueFromPipeline=$true)]
       $Second
    )
    BEGIN {
       [string[]] $p1 = $First | gm -type Properties | select -expand Name
    }
    Process {
       $Output = $First | Select $p1
       foreach($p in $Second | gm -type Properties | Where { $p1 -notcontains $_.Name } | select -expand Name) {
          Add-Member -in $Output -type NoteProperty -name $p -value $Second."$p"
       }
       $Output
    }
}

Function Export-HashtoCSV {

<#
.Synopsis
Export a hashtable to a CSV file
.Description
This function will export a hash table to a CSV file. The function
will add a new column, Type, which will be the .NET type of each
value. This information is used with Import-CSVtoHash to properly
reconstitute the hash table.

.Parameter Path
The file name and path to the CSV file.

.Parameter Hashtable
The hash table object to export.

.Example
PS C:\> $myhash | Export-HashtoCSV MyHash.csv
PS C:\> get-contnet .\Myhash.csv
#TYPE Selected.System.Collections.DictionaryEntry
"Key","Value","Type"
"name","jeff","String"
"pi","3.14","Double"
"date","2/2/2012 8:53:58 AM","DateTime"
"size","3","Int32"

.Link
Import-CSVtoHash
Export-CSV

.Inputs
Hashtable
.Outputs
None

#>

#[cmdletbinding(SupportsShouldProcess=$True)]

Param (
[Parameter(Position=0,Mandatory=$True,HelpMessage="Enter a filename and path for the CSV file")]
[ValidateNotNullorEmpty()]
[string]$Path,
[Parameter(Position=1,Mandatory=$True,HelpMessage="Enter a hashtable",
ValueFromPipeline=$True)]
[ValidateNotNullorEmpty()]
[hashtable]$Hashtable

)

Begin {
    Write-Verbose "Starting hashtable export"
    Write-Verbose "Exporting to $path"
}

Process {
    <#
      Add a column for the data type of each hash table entry.
      This can be used on import to properly reconstitute the
      hash table
    #>
    $Hashtable.GetEnumerator() | 
    Select Key,Value,@{Name="Type";Expression={$_.value.gettype().name}} | 
    Export-Csv -Path $Path

}

End {
    Write-Verbose "Ending hashtable export"
}

} #end function

Function Import-CSVtoHash {

<#
.Synopsis
Import a CSV file and create a hash table
.Description
This function will import a CSV file of hash table data and recreate
the hash table object. Ideally the CSV file will have been created
with the Export-HashtoCSV function, bu you can import any CSV provided
it has Key and Value headings. If you include a Type heading, then the
values will be cast to that type.

"Key","Value","Type"
"name","jeff","String"
"pi","3.14","Double"
"date","2/2/2012 8:53:58 AM","DateTime"
"size","3","Int32"

.Parameter Path
The file name and path to the CSV file.

.Example
PS C:\> $h=Import-CSVtoHash MyHash.csv
PS C:\> $h
Name                           Value
----                           -----
name                           jeff
pi                             3.14
date                           2/2/2012 8:53:58 AM
size                           3

.Link
Export-HashtoCSV
Import-CSV

.Inputs
String
.Outputs
Hashtable

#>

#[cmdletbinding()]

Param (
[Parameter(Position=0,Mandatory=$True,HelpMessage="Enter a filename and path for the CSV file")]
[ValidateNotNullorEmpty()]
[ValidateScript({Test-Path -Path $_})]
[string]$Path
)

Write-Verbose "Importing data from $Path"

Import-Csv -Path $path | ForEach-Object -begin {
     #define an empty hash table
     $hash=@{}
    } -process {
       <#
       if there is a type column, then add the entry as that type
       otherwise we'll treat it as a string
       #>
       if ($_.Type) {
         
         $type=[type]"$($_.type)"
       }
       else {
         $type=[type]"string"
       }
       Write-Verbose "Adding $($_.key)"
       Write-Verbose "Setting type to $type"
       
       $hash.Add($_.Key,($($_.Value) -as $type))

    } -end {
      #write hash to the pipeline
      Write-Output $hash
    }

write-verbose "Import complete"

} #end function

#######################################################################################
#  Main
#######################################################################################

$pwd = "f:\temp\don"
$messages = @()

$computers = ('UMCS20', 'UMCS19')
$Timestamps = @{}
$TimestampsFile = 'Timestamps.csv'
$outputT = get-date -format yyyyMMdd-HHmmss

$a = "<style>"
$a = $a + "BODY{background-color:peachpuff;}"
$a = $a + "TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse; padding:25px;}"
$a = $a + "TH{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color:thistle}"
$a = $a + "TD{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color:palegoldenrod}"
$a = $a + "</style>"

###
# Load back last scan time for each machine
###

if (Test-Path "$pwd\$TimestampsFile") {
	#$Timestamps = Import-CSVtoHash "$pwd\$TimestampsFile"
	Import-Csv -Path "$pwd\$TimestampsFile" | ForEach-Object `
       $Timestamps.Add($_.Key, [datetime]::ParseExact($($_.Value), "yyyyMMdd-HHmmss", $null))
       #$Timestamps[$_.Key] = [datetime]::ParseExact($($_.Value), "yyyyMMdd-HHmmss", $null))
}

###
# Check each machine
###

      #$hash = @{            
      #  Server           = ""                 
      #  Datetime         = $(get-date -format yyyyMMdd-HHmmss)
      #} 
      #$Timestamps = New-Object PSObject -Property $hash
      $Timestamps = @{}


Foreach ($Machine in $computers) {
#For ($i=1; $i -le 500; $i++)
#   $Machine = 'UMCVW' + (70000 + $i)
   write-host $Machine
   
   $cuttime = [datetime]::ParseExact("19000101-000000", "yyyyMMdd-HHmmss", $null)
   $cuttime = [datetime]::ParseExact("20120815-000000", "yyyyMMdd-HHmmss", $null)
   if ($Timestamps.ContainsKey($Machine)) {
      $cuttime = $Timestamps.Get_Item($Machine)
   } else {
      #$TimeMark | Add-Member -MemberType NoteProperty -Name Server -Value $Machine
      #$TimeMark | Add-Member -MemberType NoteProperty -Name Datetime -Value $(Get-Date)
      $Timestamps["$Machine"] = $(get-date -format yyyyMMdd-HHmmss))
      $Timestamps
   }
   Write-Host 'After ' + $cuttime
   $message = Get-EventLog -Logname System -computer $Machine -After $cuttime | `
              Where-Object {$_.EventId -eq 7 -or $_.EventId -eq 14 -or $_.EventId -eq 41 `
              							 -or $_.EventId -eq 23 -or $_.EventId -eq 110 `
              							 -or $_.EventId -eq 1117 -or $_.EventId -eq 6072 -or $_.EventId -eq 6013} 
   Write-Host $message.count   							 
   #ConvertTo-XML -NoTypeInformation).Save("$pwd\$Machine-" + $(get-date -format yyyyMMdd-HHmmss) + ".xml")
   #ConvertTo-HTML -head $a -body "<H2>Suspicious Events</H2>" | `
   #$message | ConvertTo-HTML -head $a -body "<H2>Suspicious Events</H2>" | `
   #Tee-Object $("$pwd\$Machine-" + $(get-date -format yyyyMMdd-HHmmss) + ".html")
   #Tee-Object $("f:\temp\don\" + $Machine + "-" + $(get-date -format yyyyMMdd-HHmmss) + ".xml")
   #Out-File -encoding unicode $("f:\temp\don\" + $Machine + "-" + $(get-date -format yyyyMMdd-HHmmss) + ".html")
   #Format-List | Out-File -encoding unicode $("f:\temp\don\" + $(get-date -format yyyyMMdd-HHmmss) + ".txt")

   $messages += $message

}
   
#Microsoft-Windows-Security-Auditing
#Security-Auditing

#Get-EventLog System | Where-Object {$_.EventId -eq 7001 -or $_.EventId -eq 1085} | Select-Object InstanceId,TimeGenerated,Message | Format-List | Out-File -encoding unicode $("f:\temp\don\" + $(get-date -format yyyyMMdd-HHmmss) + ".txt")
#Get-EventLog Application -EntryType Error -After (Get-Date).AddDays(-7)

# http://goo.gl/MidbQ
#4778  SESSION_RECONNECTED
#4779  SESSION_DISCONNECTED
#4800  WORKSTATION_LOCKED
#4801  WORKSTATION_UNLOCKED
#4802  SCREENSAVER_INVOKED
#4803  SCREENSAVER_DISMISSED

($messages | ConvertTo-XML -NoTypeInformation).Save("$pwd\Filter-$outputT.xml")
$Timestamps.GetEnumerator() | Select Key,Value,@{Name="Type";Expression={$_.value.gettype().name}} | Export-CSV "$pwd\$TimestampsFile"

$emailFrom = "donal_sun@umc.com"
$emailTo = "donal_sun@umc.com"
$subject = "Suspicious Events"
#$messages = Get-ChildItem UMC*.html -name $workpath | foreach ($a in $vara) {Get-Content $files -Delimiter ([char]0)}
$smtpServer = "202.14.12.4"

Send-MailMessage -To $emailTo -Subject $subject -Body " " `
								 -SmtpServer $smtpServer -From $emailFrom -Attachment "$pwd\Filter-$outputT.xml" `
								 -DeliveryNotificationOption OnFailure 

#$smtp = new-object Net.Mail.SmtpClient($smtpServer)
#$smtp.Send($emailFrom, $emailTo, $subject, $body)

#$message = New-Object System.Net.Mail.MailMessage `
#						  –ArgumentList $emailFrom, $emailTo, $subject, $body
#$attachment = New-Object System.Net.Mail.Attachment –ArgumentList 'c:\docs\test.xls', 'Application/Octet'
#$message.Attachments.Add($attachment)
#$smtp = New-Object System.Net.Mail.SMTPClient –ArgumentList 10.10.10.15
#$smtp.Send($message)
#
#function sendMail{
#
#     Write-Host "Sending Email"
#
#     #SMTP server name
#     $smtpServer = "smtp.xxxx.com"
#
#     #Creating a Mail object
#     $msg = new-object Net.Mail.MailMessage
#
#     #Creating SMTP server object
#     $smtp = new-object Net.Mail.SmtpClient($smtpServer)
#
#     #Email structure
#     $msg.From = "fromID@xxxx.com"
#     $msg.ReplyTo = "replyto@xxxx.com"
#     $msg.To.Add("toID@xxxx.com")
#     $msg.subject = "My Subject"
#     $msg.body = "This is the email Body."
#
#     #Sending email
#     $smtp.Send($msg)
# 
#}
#
##Calling function
#sendMail

function Merge-XmlFile
{
<#
.SYNOPSIS
Merge source XML file to target XML file and update only modified element. 
		
.DESCRIPTION
The function takes Xml filepath and merge the content to Xml filepath output. It is usuful to merge any 
XML's like file, such as .proj, web.config, app.config.
In that case you will be able to package only the essential element, and this function will merge it 
automatically. The old Xml target file will be stored as a backup file.
    
.INPUTS
None. You can not pipe objects to Merge-XmlFile
	
.OUTPUTS
None. No output from Merge-XmlFile
		
.PARAMETER sourceXmlFile
MANDATORY parameter to specify the source Xml filename. 
	
.PARAMETER targetXmlFile
MANDATORY parameter to specify target Xml filename.
        	
.EXAMPLE
    PS>  Merge-XmlElement "additionalweb.config" "web.config"
			

	Description
	-----------
	Update web.config based on additionalweb.config content. The function will search and update according the the tag,
    attribute and element.
    When there is doubt in the Xml, you have to specify Select keyword to select specific key, or
    Remove keyword to remove specific key. For example appSettings section in the web.config, as follows
    
    Example: 
    additionalweb.config
    <configuration>
       <appSettings>
         <!--Select=add[@key='Keyword']-->
         <add key="Keyword" value="SharePoint,PowerShell" />
         <!--Remove=add[@key='OldKeyword']-->
         <add key="OldKeyword" value="SharePoint 2007" />
       </appSettings>
    </configuration>
    
    web.config
    <configuration>
       <appSettings>
         <add key="Keyword" value="SharePoint 2007,PowerShell" />         
         <add key="OldKeyword" value="SharePoint 2007" />
       </appSettings>
    </configuration>
    
    After the operation, web.config will become
    <configuration>
       <appSettings>
         <add key="Keyword" value="SharePoint,PowerShell" />                  
       </appSettings>
    </configuration>
    
		            
.LINK
    Author blog  : IdeasForFree  (http://blog.libinuko.com)
.LINK
    Author email : cakriwut@gmail.com
#>
   param ( 
        [Parameter(Mandatory=$true,Position=0)]            
        $sourceXmlFile, 
        [Parameter(Mandatory=$true,Position=1)]       
        $targetXmlfile 
  )
  
  if(!(test-path $sourceXmlFile))
  {
    write-host "Can not find source XML file. $sourceXmlFile."
  }
  if(!(test-path $targetXmlFile))
  {
    write-host "Can not find target XML file. $targetXmlFile."
  }
   
   $target = gi $gargetXmlFile
   $backup = (join-path $target.Directory $target.BaseName) + "_" + (get-date).tostring("yyyy_MM_dd_hh_mm_ss") + ".bak"  
  
   $xmlSource = [xml](get-content $sourceXmlFile)  
   $xmlTarget = [xml](get-content $targetXmlfile)
   #save backup
   $xmlTarget.Save($backup)
   
   $SourceElement = $xmlSource.get_Documentelement()
   $TargetElement = $xmlTarget.get_DocumentElement() 
  
   Merge-XmlElement $SourceElement $TargetElement
   
   #save backup
   $xmlTarget.Save($targetXmlFile)
}

function Merge-XmlElement 
{ 
<#
.SYNOPSIS
Merge source XML element to target XML element and update only modified element. 
		
.DESCRIPTION
The function takes XmlElement input and merge the content to XmlElement output. It is usuful to merge any XML's like file, such as .proj, web.config, app.config.
In that case you will be able to package only the essential element, and this function will merge it automatically.
    
.INPUTS
None. You can not pipe objects to Merge-XmlElement
	
.OUTPUTS
None. No output from Merge-XmlElement
		
.PARAMETER sourceElement
MANDATORY parameter to specify the source Xml element. 
	
.PARAMETER targetElement
MANDATORY parameter to specify target Xml element.
        	
.EXAMPLE
    PS>  $xmlWebConfig = [xml](get-content "web.config")  
    PS>  $xmlUpdateWebConfig = [xml](get-content "additionalweb.config")
	PS>  $targetRoot = $xmlWebConfig.get_DocumentElement()
    PS>  $sourceRoot = $xmlUpdateWebConfig.get_DocumentElement()
    PS>  Merge-XmlElement $sourceRoot $targetRoot
			

	Description
	-----------
	Update web.config based on additionalweb.config content. The function will search and update according the the tag, attribute and element.
    When there is doubt in the Xml, for example appSettings section in the web.config, you have to specify Select keyword to select specific key, or
    Remove keyword to remove specific key.
    
    Example: 
    additionalweb.config
    <configuration>
       <appSettings>
         <!--Select=add[@key='Keyword']-->
         <add key="Keyword" value="SharePoint,PowerShell" />
         <!--Remove=add[@key='OldKeyword']-->
         <add key="OldKeyword" value="SharePoint 2007" />
       </appSettings>
    </configuration>
    
    web.config
    <configuration>
       <appSettings>
         <add key="Keyword" value="SharePoint 2007,PowerShell" />         
         <add key="OldKeyword" value="SharePoint 2007" />
       </appSettings>
    </configuration>
    
    After the operation, web.config will become
    <configuration>
       <appSettings>
         <add key="Keyword" value="SharePoint,PowerShell" />                  
       </appSettings>
    </configuration>
    
		            
.LINK
    Author blog  : IdeasForFree  (http://blog.libinuko.com)
.LINK
    Author email : cakriwut@gmail.com
#>
  param ( 
        [Parameter(Mandatory=$true,Position=0)]
        [System.Xml.XmlElement]                    
        $sourceElement,
        [Parameter(Mandatory=$true,Position=1)]
        [System.Xml.XmlElement]        
        $targetElement 
  )
    
  if ($sourceElement.get_Name() -ne $targetElement.get_Name()) 
  { 
    write-host "Source element name $($sourceElement.get_Name()) and target element name $($targetElement.get_Name()) do not match" 
    return
  } 
     
  if (-not $sourceElement.get_HasChildNodes()) { return } 
  
  $sourceChildren = $sourceElement.get_Childnodes() 
  $targetChildren = $targetElement.get_Childnodes()
  $prevChild = $null
  
  foreach ($sourceChild in $sourceChildren) 
  {     
     if ($sourceChild.get_Name() -eq "#comment") 
     { 
       $prevChild = $sourceChild       
       continue 
     }
              
     $matchingNode = $False 
     $targetChild = $Null 
     $select = $False
     $remove = $False
     
     foreach ($child in $targetChildren )
     {
        $targetChild = $child
        if(($select = ($prevChild -and $prevChild.Value.StartsWith("Select="))) `
            -or ($remove = ($prevChild -and $prevChild.Value.StartsWith("Remove=")))) { break; }        
                            
        if ($sourceChild.get_Name() -eq $targetChild.get_Name())
        {         
            OverrideAttribute $sourceChild $targetChild
            $matchingNode = $True
            break;       
         }       
     } #end foreach TargetChildren
                 
     if ($matchingNode -eq $False) 
     { 
        if($select)
        {            
           if(($selectedElement = $targetElement.SelectSingleNode($prevChild.Value.Trim().Remove(0,7))))
           {
              OverrideAttribute $sourceChild $selectedElement
           } else {
              AppendElement $sourceChild $targetElement
           }
         } elseif($remove)
         {
           if(($selectedElement = $targetElement.SelectSingleNode($prevChild.Value.Trim().Remove(0,7))))
           {
              write-host "Removing element " $selectedElement.OuterXml
              $selectedElement.RemoveAll()
           }
         } else 
         {
            AppendElement $sourceChild $targetElement
         }

      } else { 
         if($sourceChild.get_HasChildNodes()) {
            Merge-XmlElement $sourceChild $targetChild
             }
      }               
      $prevChild = $null
   } #end foreach SourceChildren
   
}

function OverrideAttribute
{
   param(
      $Source,
      $Target
   )
   
   foreach($SourceAttribute in $Source.get_Attributes())
   {
      if($SourceAttribute.get_Value() -ne $Target.GetAttribute($SourceAttribute.get_Name())) 
      {
          write-host "Override attribute " $SourceAttribute.get_Name() "," $Target.GetAttribute($SourceAttribute.get_Name()) "=>" $SourceAttribute.get_Value() 
      }
      $Target.SetAttribute($SourceAttribute.get_Name(),$SourceAttribute.get_Value())
   }
}

function AppendElement
{
   param (
       $Source,
       $Target
   )
   $NewElement = $Source.CloneNode($True)
   $Target.AppendChild($Target.get_OwnerDocument().ImportNode($NewElement,$True))
   $Target.Normalize()
}

