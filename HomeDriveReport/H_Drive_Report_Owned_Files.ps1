#Requires -Modules ImportExcel -version 2
<#
.SYNOPSIS
  Queries Identity Finder SQL database for results and  generates files/emails for review.
.DESCRIPTION
  Script is compatible with Powershell version 3.0+, and assumes a CIFS fileshare.
  To make compatible with PS version 2, remove [ordered] type constraint on lines 348 and 403
  The OutputFormat parameter is required.
  The EmailTarget parameter defaults to "none" and writes info to log.
  The default file name will be the value of $resultFileName.
  Contact Spirion Support or your salesperson to get table names.
.PARAMETER OutputFormat
  This parameter allows you to choose between HTML and Excel for results file output. 
  You can also choose to not write files to disk. This option will instead write a log
  message containing the username of the person whose results are being processed. 
.PARAMETER EmailTarget
  This parameter allows you to choose a recipient for notification emails. 
  Choosing 'Administrator' will send all result emails to a single mailbox.
  Choosing 'User' will send an email to each individual who had results.
  Choosing 'None' will write a log message containing the username of the individual 
  whose results are being processed.
.PARAMETER NamedFiles
  The NamedFiles switch allows you to name each file with the username of the 
  individual whose results are being processed. This is useful if you want to 
  save all the results to a single location. If NamedFiles is not enabled, 
  all files will have the name specified in $resultFileName.
.INPUTS
  None
.OUTPUTS
  HTML or Excel (ImportExcel module: https://github.com/dfinke/ImportExcel)
.NOTES
  Version:        1.5
  Author:         dd4495 
  Creation Date:  October 2016
#>

    Param 
    ( 
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)][ValidateSet('xlsx', 'html', 'csv', 'None')]
        [string]$OutputFormat,
        [Parameter(ValueFromPipelineByPropertyName=$true)][ValidateSet('None', 'User', 'Administrator')]
        [string]$EmailTarget='None',
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [switch]$NamedFiles
        )
#---------------------------------------------------------[Initializations]--------------------------------------------------------
$fOwners=@()
#----------------------------------------------------------[Declarations]----------------------------------------------------------
#SQL Server setup
$SQLServerFQDN = 'consoleserver.starfleet.edu'
$SQLDatabaseName = 'IdfMC'
$domain = 'starfleet'
#Find your endpoint ID by running: 
# select * from [$table0] 
#    where name='$ShortServerName'; 
$endpoint = '-2147483603'

#Table Names
$d = '[$table1]'
$l = '[$table2]'
$m = '[$table3]'
$h = '[$table4]'

#Report configuration
$backgroundColor = '#6495ED'
$oddRowColor = '#ffffff'
$evenRowColor = '#dddddd'
$Title = 'Report Title'

#Email Setup
$Script:mailServer = 'mail.starfleet.edu'
$emailDomain = 'starfleet.edu'
$NotificationSubject = 'Identity Finder Scan Results'
$idfAdmin = 'identityfinder'
$testadmin = 'admin'

#Results Setup
#To save files to mapped network share, set $resultFilePath to mapped drive letter
$Script:prettyDate = Get-Date -Format MM-dd-yyyy
$Script:date = Get-Date -Format ddMMyy_HHmmssfff
$resultFilePath = 'C:\temp'
$resultFileName = ('IdentityFinderResults-{0}.{1}' -f $date, $OutputFormat)

#Big results setup
$BigResult = 5000
$BigNotificationSubject = 'IDF Large Results'
$bigResultFile = 'C:\temp\bigowners.txt'

#Logging Setup
$logPath = ('C:\temp\idfLog-{0}.log' -f $date)
#-----------------------------------------------------------[Functions]------------------------------------------------------------

Function Write-Log { 
<# 
.Synopsis 
   Write-Log writes a message to a specified log file with the current time stamp. 
.DESCRIPTION 
   The Write-Log function is designed to add logging capability to other scripts. 
   In addition to writing output and/or verbose you can write to a log file for 
   later debugging. 
.NOTES 
   Created by: Jason Wasser @wasserja 
   Modified: 11/24/2015 09:30:19 AM   
 
   Changelog: 
    * Code simplification and clarification - thanks to @juneb_get_help 
    * Added documentation. 
    * Renamed LogPath parameter to Path to keep it standard - thanks to @JeffHicks 
    * Revised the Force switch to work as it should - thanks to @JeffHicks 
 
   To Do: 
    * Add error handling if trying to create a log file in a inaccessible location. 
    * Add ability to write $Message to $Verbose or $Error pipelines to eliminate 
      duplicates. 
.PARAMETER Message 
   Message is the content that you wish to add to the log file.  
.PARAMETER Path 
   The path to the log file to which you would like to write. By default the function will  
   create the path and file if it does not exist.  
.PARAMETER Level 
   Specify the criticality of the log information being written to the log (i.e. Error, Warning, Informational) 
.PARAMETER NoClobber 
   Use NoClobber if you do not wish to overwrite an existing file. 
.EXAMPLE 
   Write-Log -Message 'Log message'  
   Writes the message to c:\Logs\PowerShellLog.log. 
.EXAMPLE 
   Write-Log -Message 'Restarting Server.' -Path c:\Logs\Scriptoutput.log 
   Writes the content to the specified log file and creates the path and file specified.  
.EXAMPLE 
   Write-Log -Message 'Folder does not exist.' -Path c:\Logs\Script.log -Level Error 
   Writes the message to the specified log file as an error message, and writes the message to the error pipeline. 
.LINK 
   https://gallery.technet.microsoft.com/scriptcenter/Write-Log-PowerShell-999c32d0 
#> 
    Param 
    ( 
        [Parameter(Mandatory=$true,HelpMessage='Please add a log message.', ValueFromPipelineByPropertyName=$true)][ValidateNotNullOrEmpty()][Alias('LogContent')] 
        [string]$Message,
        [Parameter(ValueFromPipelineByPropertyName=$true)][Alias('LogPath')] 
        [string]$Path='c:/temp/PowerShellLog.log', 
        [Parameter(ValueFromPipelineByPropertyName=$true)][ValidateSet('Error','Warn','Info')] 
        [string]$Level='Info', 
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [switch]$NoClobber 
    ) 
 
    Begin { 
        # Set VerbosePreference to Continue so that verbose messages are displayed. 
        $VerbosePreference = 'Continue' 
        } 
    Process {
        # If the file already exists and NoClobber was specified, do not write to the log. 
        if ((Test-Path -Path $Path) -AND $NoClobber) { 
            Write-Error -Message ('Log file {0} already exists, and you specified NoClobber. Either delete the file or specify a different name.' -f $Path) 
            Return 
            }
        # If attempting to write to a log file in a folder/path that doesn't exist create the file including the path. 
        elseif (!(Test-Path -Path $Path)) { 
            Write-Verbose -Message ('Creating {0}.' -f $Path) 
            $NewLogFile = New-Item -Path $Path -Force -ItemType File 
            } 
        else { 
            # Nothing to see here yet. 
            } 
 
        # Format Date for our Log File 
        $FormattedDate = Get-Date -Format 'yyyy-MM-dd HH:mm:ss' 
 
        # Write message to error, warning, or verbose pipeline and specify $LevelText 
        switch ($Level) { 
            'Error' { 
                Write-Error -Message $Message
                $LevelText = 'ERROR:' 
                } 
            'Warn' { 
                Write-Warning -Message $Message 
                $LevelText = 'WARNING:' 
                } 
            'Info' { 
                Write-Verbose -Message $Message 
                $LevelText = 'INFO:' 
                } 
            } 
         
        # Write log entry to $Path 
        ('{0} {1} {2}' -f $FormattedDate, $LevelText, $Message) | Out-File -FilePath $Path -Append 
    } 
    End { 
    } 
}
Function Send-ResultMail {
<#
.SYNOPSIS
  Sends an email to the specified email address
.DESCRIPTION
  This function sends an email to the specified address. 
  Subject and From fields have defaults that can be overridden
.EXAMPLE
  Send-ResultMail -To 'thor@asgard.edu' -From 'loki@asgard.edu'
.EXAMPLE
  Send-ResultMail -To 'thor@asgard.edu -From 'loki@asgard.edu' -Subject 'Tesseract'
.PARAMETER Subject
  The subject of the email. Defaults to 'Test Message'
.PARAMETER Attachment
  The file to attach to the email. This parameter is not required.
.PARAMETER To
  The destination email address. This parameter is required.
.PARAMETER From
  The source email address. This parameter is required.
.NOTES
    Version:        1.2
    Author:         dd4495 
    Creation Date:  August 2015
#>

    [CmdletBinding(SupportsShouldProcess=$true,ConfirmImpact='Low')]
    param (
        [Parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,Position=0)]
        [string]$Subject = 'Test Message',
        [Parameter(Mandatory=$true,HelpMessage='Please add a "To" address',ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,Position=1)]
        [string]$To,
        [Parameter(Mandatory=$true,ValueFromPipeline=$true,HelpMessage='Please add a "From" address',ValueFromPipelineByPropertyName=$true,Position=2)]
        [string]$From,
        [Parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,Position=3)]
        [array]$Body = $Body,
        [Parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,Position=4)]
        [string]$Attachment
    )

    $smtpServer = $mailServer
    $Message = New-Object -TypeName System.Net.Mail.MailMessage
    $Message.From = $From
    $Message.To.Add($To)
    $Message.Subject = $Subject
    $Message.IsBodyHtml = $True
    $Message.Body = $Body
    if ($Attachment){$Message.Attachments.Add($Attachment)}
    $SMTP = New-Object -TypeName Net.Mail.SmtpClient -ArgumentList ($smtpServer)
    $SMTP.Send($Message)
    
}
#-----------------------------------------------------------[Execution]------------------------------------------------------------

#region first SQL Query
$SqlQuery =(@'
select distinct d.FileOwner
	from {0} d
	join {1} l on (d.MatchLocationId = l.Id)
			where l.EndpointId={2}
            order by d.FileOwner asc;
'@ -f $d, $l, $endpoint)

$SqlConnection = New-Object -TypeName System.Data.SqlClient.SqlConnection 
$SqlConnection.ConnectionString = ('Server = {0}; Database = {1}; Integrated Security = True' -f $SQLServerFQDN, $SQLDatabaseName)

$SqlCmd = New-Object -TypeName System.Data.SqlClient.SqlCommand
$SqlCmd.CommandText = $SqlQuery
try {$SqlCmd.Connection = $SqlConnection
    Write-Log -Message 'Creating SQL connection' -Level Info -Path $logPath}
catch {Write-Log -Message "Couldn't connect to SQL server. Do you have a SQL adapter?" -Level Error -Path $logPath}

$SqlAdapter = New-Object -TypeName System.Data.SqlClient.SqlDataAdapter
$SqlAdapter.SelectCommand = $SqlCmd

$DataSet = New-Object -TypeName System.Data.DataSet
$null = $SqlAdapter.Fill($DataSet)

$SqlConnection.Close()

$data = $DataSet.Tables[0]
#endregion

#Get unique file owners so each can be queried individually
for ($i=0; $i -lt $data.Rows.Count; $i++) {
    $fOwners += $data.Rows[$i].FileOwner
}
$Owners = $fOwners | Select-Object -Unique

#Loop through owners in $Owners array and query SQL for all results matching that owner
#The "if ($owner -match $domain)" statement will filter out any results that are
# not owned by a domain user/domain group (e.g. files owned by BUILTIN/Administrators
# will be skipped)
foreach ($owner in $owners) {
    if ($owner -match $domain){
    #region Second SQL Query
        $SqlQuery =(@"
        select distinct l.Location, d.FileOwner, l.InternalLocationType, dbo.conv_IdentityType(m.MatchTypeId)AS MatchType, h.Count AS Hits
            from {0} l
	        join {1} d on (d.MatchLocationId = l.Id)
		    join {2} m on (m.matchlocationid = d.matchlocationid)
		    join {3} h on (m.Id = h.Id)
		        where  l.EndpointId = {4}
                and d.FileOwner LIKE '{5}';
"@ -f $l, $d, $m, $h, $endpoint, $owner)

        $SqlConnection = New-Object -TypeName System.Data.SqlClient.SqlConnection
        $SqlConnection.ConnectionString = ('Server = {0}; Database = {1}; Integrated Security = True' -f $SQLServerFQDN, $SQLDatabaseName)

        $SqlCmd = New-Object -TypeName System.Data.SqlClient.SqlCommand
        $SqlCmd.CommandText = $SqlQuery
        $SqlCmd.Connection = $SqlConnection

        $SqlAdapter = New-Object -TypeName System.Data.SqlClient.SqlDataAdapter
        $SqlAdapter.SelectCommand = $SqlCmd

        $DataSet = New-Object -TypeName System.Data.DataSet
        $null = $SqlAdapter.Fill($DataSet)

        $SqlConnection.Close()

        $data = $DataSet.Tables[0]
        #endregion

        #Declare arrays so they'll be reset with each loop
        $Location=@()
        $Owner=@()
        $LocType=@()
        $MatType=@()
        $Hits=@()
        
        #Count the number of results an owner has.
        #If >= $BigResult, owner will be skipped and their name will be sent to an administrator
        $intCount = $Data.Rows.Count
        $BigOwners = @()

        #Format output. Can be customized using variables at beginning of script.
        $Header = (@'
        <style>
        TABLE {{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}}
        TH {{border-width: 1px;padding: 3px;border-style: solid;border-color: black;background-color: {0};}}
        TD {{border-width: 1px;padding: 3px;border-style: solid;border-color: black;}}
        .odd  {{ background-color:{1}; }}
        .even {{ background-color:{2}; }}
        </style>
        <title>
        {3}
        </title>
'@ -f $backgroundColor, $oddRowColor, $evenRowColor, $Title)

    for ($i=0; $i -lt $data.Rows.Count; $i++) {
        $Location += $data.Rows[$i].Location
        $Owner += $data.Rows[$i].FileOwner
        $LocType += $data.Rows[$i].InternalLocationType
        $MatType += $data.Rows[$i].MatchType
        $Hits += $data.Rows[$i].Hits

        $properties = [ordered]@{'Owner'          = $Owner;
                                'Location'        = $Location;
                                'LocationType'    = $LocType;
                                'MatchType'       = $MatType;
                                'Hits'            = $Hits;
                                }
        }

    $object = New-Object -TypeName psobject -Property $properties

    if ($intCount -gt $BigResult) {
        $BigOwner = $object.owner[0]
        Write-Log -Message "$BigOwner has over 5000 results." -Level Info -Path $logPath
        $BigOwners += $BigOwner
        $BigOwners >> $bigResultFile
    }
    
    else {
    $userarray = @() 

        for($k=0; $k -lt $object.Owner.Count; $k++) {      
            $own = $object.owner[$k]
            $match = $object.MatchType[$k]
            $loct = $object.LocationType[$k]
            $loc = $object.Location[$k]
            $hts = $object.Hits[$k]
            
            #Replaces SQL location code with human-readable equivalent
            switch ($loct){
                0  {$loct = 'None'}
                1  {$loct = 'Windows Mail or Outlook Express E-Mail Message'}
                2  {$loct = 'Windows Mail or Outlook Express Attachment'}
                3  {$loct = 'Outlook E-Mail Message'}
                4  {$loct = 'Outlook Attachment'}
                5  {$loct = 'Internet Explorer Browser Data'}
                6  {$loct = 'Firefox Browser Data'}
                7  {$loct = 'Windows Registry'}
                8  {$loct = 'Compressed File'}
                9  {$loct = 'File'}
                10 {$loct = 'Rec Moved'}
                11 {$loct = 'Mdb File'}
                12 {$loct = 'Web Page'}
                13 {$loct = 'Database Table'}
                14 {$loct = 'Thunderbird E-Mail Message'}
                15 {$loct = 'Thunderbird Attachment'}
                16 {$loct = 'MBOX File E-Mail Message'}
                17 {$loct = 'MBOX Attachment'}
                18 {$loct = 'Lotus Notes E-Mail Message'}
                19 {$loct = 'Lotus Notes Attachment'}
                20 {$loct = 'Microsoft Exchange E-Mail Message'}
                21 {$loct = 'Microsoft Exchange Attachment'}
                22 {$loct = 'Shadow File'}
                23 {$loct = 'SharePoint'}
    }
      
            $props = [ordered]@{'Owner'             = $own;
                                'Location'          = $loc;
                                'LocationType'      = $loct;
                                'MatchType'         = $match;
                                'Hits'              = $hts
                    }

            $userarrayinfo = New-Object -TypeName psObject -Property $props
            $oName = $userarrayinfo.Owner -split '\\' | Select-Object -Index 1
            $count = $object.owner.count
            $userarray += $userarrayinfo
            $Body = (@"
            <!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
            <html xmlns="http://www.w3.org/1999/xhtml">
                <head><title>Identity Finder Report</title></head>
                <body>
                    <p style="padding:10; margin-bottom:10; margin-top:10; font-size:20px;">
                        Identity Finder has discovered {0} potential PII results in H:\{1}.
                        <br> 
                        To review your results, open {2} in your H drive.
                    </p>
    
                    <br><br>
    
                    <p align="center" style="font-family:'Lucida Sans Unicode', 'Lucida Grande', sans-serif; font-size:10px; line-height:14px; color:#646464;">
                        This report generated for {1} on {3} by Identity Finder.
                        <br>
                    </p>
                </body>
            </html>
"@ -f $count, $oName, $resultFileName, $prettydate)
        
            if ($OutputFormat -eq 'html') {$Output = Write-Output -InputObject $userarray | ConvertTo-Html -Head $Header}
            elseif (($OutputFormat -eq 'xlsx') -or ($OutputFormat -eq 'csv')){$Output = Write-Output -InputObject $userarray}
            else {Write-Log -Message "You might want to exorcise your computer if you're getting this error. (On or around line 437)" -Level Error -Path $logPath}
        }   
        $fname = $userarray[0].Owner -split '\\' |  Select-Object -Index 1
        #Logic for selecting whether emails will be sent to individual users or an administrator
        if ($emailTarget -eq 'User'){
            Send-ResultMail -Subject $NotificationSubject -To ('{0}@{1}' -f $fname, $emailDomain) -From ('{0}@{1}' -f $idfAdmin, $emailDomain)
        }
        elseif ($emailTarget -eq 'Administrator'){
            Send-ResultMail -Subject $NotificationSubject -To ('{0}@{1}' -f $testadmin, $emailDomain) -From ('{0}@{1}' -f $idfAdmin, $emailDomain)
        }
        else {Write-Log -Message ('No email recipient chosen. Processing results for {0} silently.' -f $fname) -Level Info -Path $logPath}

        if ($namedFiles) {$resultFileName = "$fname.$OutputFormat"}
        else {$resultFileName = $resultFileName}
        #Logic for outputting the correct filetype
        if ($OutputFormat -eq 'html') {
            $Output | Out-File -Force -FilePath ('{0}\{1}' -f $resultFilePath, $resultFileName)
        }
        elseif ($OutputFormat -eq 'xlsx') {
            $Output | Export-Excel -Path ('{0}\{1}' -f $resultFilePath, $resultFileName)
        }
        elseif ($OutputFormat -eq 'csv') {
            $Output | Export-Csv -Path ('{0}\{1}' -f $resultFilePath, $resultFileName)
        }

        else {Write-Log -Message ('Processing results for {0} (Timestamp:{1})' -f $fname, $date) -Level Info -Path $logPath}
    }
    
    }

}
#region Big Results
#Uncomment this section to notify an admin about users who have a large number of results
#Send-ResultMail -Subject $BigNotificationSubject -To ('{0}@{1}' -f $testadmin, $emailDomain) -From ('{0}@{1}' -f $idfAdmin, $emailDomain) -Attachment $bigResultFile
#endregion