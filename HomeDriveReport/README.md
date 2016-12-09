
#SYNOPSIS
  Queries Identity Finder SQL database for results and  generates files/emails for review.
#DESCRIPTION
  Script is compatible with Powershell version 3.0+, and assumes a CIFS fileshare.
  To make compatible with PS version 2, remove [ordered] type constraint on lines 348 and 403.
  The OutputFormat parameter is required.
  The EmailTarget parameter defaults to "none" and writes info to log.
  The default file name will be the value of $resultFileName.
  Contact Spirion Support or your salesperson to get table names.

#PARAMETER OutputFormat

  This parameter allows you to choose between HTML and Excel for results file output. 
  You can also choose to not write files to disk. This option will instead write a log
  message containing the username of the person whose results are being processed. 

#PARAMETER EmailTarget

  This parameter allows you to choose a recipient for notification emails. 
  Choosing 'Administrator' will send all result emails to a single mailbox.
  Choosing 'User' will send an email to each individual who had results.
  Choosing 'None' will write a log message containing the username of the individual 
  whose results are being processed.

#PARAMETER NamedFiles

  The NamedFiles switch allows you to name each file with the username of the 
  individual whose results are being processed. This is useful if you want to 
  save all the results to a single location. If NamedFiles is not enabled, 
  all files will have the name specified in $resultFileName.

#INPUTS

  None

#OUTPUTS

  HTML or Excel (ImportExcel module can be found [here](https://github.com/dfinke/ImportExcel))

#NOTES

  - Version:        1.5
  - Author:         dd4495 
  - Creation Date:  October 2016
