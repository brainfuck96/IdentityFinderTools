# IdF_ExchServerReport
Powershell script for parsing Identity Finder Exchange scan results

## Usage

* Run a remote Exchange server scan in Identity Finder and export the results to a CSV.
* Change the $CSVLocation variable to import the CSV to Powershell.
* Change the values of $mailServer, $mailTo, and $mailFrom to use your Exchange server to email the results to a single mailbox. 
* To mail each individual user their results, set the $mailTo variable equal to $fn. 


## Report

* $Title will change the title of the report
* $backgroundColor will change the color of the header
* $oddRowColor and $evenRowColor will change the row colors
