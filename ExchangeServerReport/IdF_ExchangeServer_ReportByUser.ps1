#requires -version 2
<#
.SYNOPSIS
  Parses Identity Finder Exchange scan results into user-specific tables
.DESCRIPTION
  This script will import a CSV with data from an Identity Finder results export, 
  parse the data, and return lists of locations and PII types separated by user
.PARAMETER <Parameter_Name>
    None
.INPUTS
  CSV file
.OUTPUTS
  HTML formatted emails
  log files
.NOTES
  Version:        1.0
  Author:         dd4495 
  Creation Date:  July 2016
#>

#---------------------------------------------------------[Initializations]--------------------------------------------------------

#Set Error Action to Silently Continue
$ErrorActionPreference = 'SilentlyContinue'

# Declare arrays
$location=@();$dataType=@();$locationType=@();$stringcoll = @()
$username = $null;$email = $null
$un=@();$DatType=@();$LocType=@();$loc=@()

#----------------------------------------------------------[Declarations]----------------------------------------------------------
$csvLocation = 'C:\temp\EmailSearchResults.csv'

# Mail Settings
$Script:mailServer = 'mail.asgard.edu'
$mailTo = 'admin@asgard.edu'
$mailFrom = 'identityfinder@asgard.edu'

# Report settings
$Title = 'Identity Finder Scan Results'
$backgroundColor = '#6495ED'
$oddRowColor = '#ffffff'
$evenRowColor = '#dddddd'

#-----------------------------------------------------------[Functions]------------------------------------------------------------
Function Send-ResultMail {
[CmdletBinding(SupportsShouldProcess=$True,ConfirmImpact='Low')]
param (
    [Parameter(Mandatory=$false,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True,Position=0)]
    [string]$Subject = 'Test Message',
    [Parameter(Mandatory=$True,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True,Position=1)]
    [string]$To,
    [Parameter(Mandatory=$false,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True,Position=2)]
    [string]$From = 'loki@asgard.edu',
    [Parameter(Mandatory=$false,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True,Position=3)]
    [array]$Body = $Body)

    $smtpServer = $mailServer
    $Message = New-Object System.Net.Mail.MailMessage
    $Message.From = $From
    $Message.To.Add($To)
    $Message.Subject = $Subject
    $Message.IsBodyHtml = $True
    $Message.Body = $Body
    $SMTP = New-Object Net.Mail.SmtpClient($smtpServer)
    $SMTP.Send($Message)
    
}
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
    [CmdletBinding()] 
    Param 
    ( 
        [Parameter(Mandatory=$true, 
                   ValueFromPipelineByPropertyName=$true)] 
        [ValidateNotNullOrEmpty()] 
        [Alias('LogContent')] 
        [string]$Message, 
 
        [Parameter(Mandatory=$false)] 
        [Alias('LogPath')] 
        [string]$Path='C:\Logs\PowerShellLog.log', 
         
        [Parameter(Mandatory=$false)] 
        [ValidateSet('Error','Warn','Info')] 
        [string]$Level='Info', 
         
        [Parameter(Mandatory=$false)] 
        [switch]$NoClobber 
    ) 
 
    Begin 
    { 
        # Set VerbosePreference to Continue so that verbose messages are displayed. 
        $VerbosePreference = 'Continue' 
    } 
    Process 
    { 
         
        # If the file already exists and NoClobber was specified, do not write to the log. 
        if ((Test-Path $Path) -AND $NoClobber) { 
            Write-Error "Log file $Path already exists, and you specified NoClobber. Either delete the file or specify a different name." 
            Return 
            } 
 
        # If attempting to write to a log file in a folder/path that doesn't exist create the file including the path. 
        elseif (!(Test-Path $Path)) { 
            Write-Verbose "Creating $Path." 
            $NewLogFile = New-Item $Path -Force -ItemType File 
            } 
 
        else { 
            # Nothing to see here yet. 
            } 
 
        # Format Date for our Log File 
        $FormattedDate = Get-Date -Format 'yyyy-MM-dd HH:mm:ss' 
 
        # Write message to error, warning, or verbose pipeline and specify $LevelText 
        switch ($Level) { 
            'Error' { 
                Write-Error $Message 
                $LevelText = 'ERROR:' 
                } 
            'Warn' { 
                Write-Warning $Message 
                $LevelText = 'WARNING:' 
                } 
            'Info' { 
                Write-Verbose $Message 
                $LevelText = 'INFO:' 
                } 
            } 
         
        # Write log entry to $Path 
        "$FormattedDate $LevelText $Message" | Out-File -FilePath $Path -Append 
    } 
    End 
    { 
    } 
}
#-----------------------------------------------------------[Execution]------------------------------------------------------------

#region OtherVariables
try {$table = Import-Csv -Path $csvLocation -Delimiter ',' -ErrorAction Stop; Write-Log -Message 'Imported CSV file' -Level Info} 
catch {write-log -message 'Import Failed, please check your path and try again' -level Error} 

$Header = @"
<style>
TABLE {border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
TH {border-width: 1px;padding: 3px;border-style: solid;border-color: black;background-color: $backgroundColor;}
TD {border-width: 1px;padding: 3px;border-style: solid;border-color: black;}
.odd  { background-color:$oddRowColor; }
.even { background-color:$evenRowColor; }
</style>
<title>
$Title
</title>
"@
#endregion

#region setup
# Pull individual columns from the CSV variable
for ($i=0;$i -lt $table.Count;$i++) {
    foreach ($val in $table[$i]) {
        $location += $table[$i].Location
        $dataType += $table[$i].'Data Type'
        $locationType += $table[$i].'Location Type'
    }
}

# Get rid of all text prior to the username in the Locations column
foreach ($string in $location) {
    $nustring = $string -split ':',3 | Select-Object -Index 2
    $stringcoll += $nustring -replace '^(.*?)\\(.*)','$1=$2' 
}

# Split the new, shorter Location column into two columns, one containing the username, the other
# containing the actual location data. Create an object containing all 4 columns.
foreach ($line in $stringcoll) {
    $username += @($line -split '=' | Select-Object -Index 0)
    $email += @($line -split '=' | Select-Object -Index 1)

    $properties = @{'Username'=$username;
                    'Location'=$email;
                    'LocationType'=$locationType;
                    'DataType'=$dataType;
                    }
}
$object = New-Object -TypeName psobject -Property $properties

# Get a list of usernames without duplicates
$uniqueUsers = $username | sort-object | Get-Unique
#endregion

#region main
# Go through the contents of the object created on line 108 and compare each username in the object
# to the list of unique usernames 
# Create a new array containing all the results for an individual user, for each user in $unique users.
# Send an HTML formatted email to specified user with results in the body of the email.
for ($j=0; $j -lt $uniqueUsers.Count; $j++) {
    $fn = $uniqueUsers[$j].TrimStart(' ')
    $userarray = @() 
    for($k=0; $k -lt $object.Username.Count; $k++) {
        if ($object.Username[$k] -match $uniqueUsers[$j]) {
            
            $un = $object.Username[$k]
            $DatType = $object.DataType[$k]
            $LocType = $object.LocationType[$k]
            $loc = $object.Location[$k]

            $props = @{'Username' = $un;
                    'Location'=$loc;
                    'LocationType'=$LocType;
                    'DataType'=$DatType;
            }
            $userarrayinfo = New-Object -TypeName psObject -Property $props
            $userarray += $userarrayinfo
            $Body = Write-Output $userarray | ConvertTo-Html -Head $Header 
        }   
    } 
    try{ Send-ResultMail -Subject 'Identity Finder Scan Results' -To $mailTo -From $mailFrom -ErrorAction Stop; Write-Log -Message "Sent email to $mailto" -Level Info}
    catch {write-log -message "Send message failed with error $Error[0]. Please try again." -level Error}
#endregion      
}