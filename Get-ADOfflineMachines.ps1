#This script is NASTY, thrown together in a hurry. Develop if used regularly.

function Get-ADOfflineComp{
    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory=$False,
            Position=0,
            HelpMessage="Enter an OU, default is OU=Servers,DC=paconsulting,DC=com"
        )]
        [string]$Searchbase = "OU=Servers,DC=paconsulting,DC=com",

        [Parameter(
            Mandatory=$False,
            Position=1,
            HelpMessage="the amount of inactive time to search for, the default is 90"
        )]
        [int]$Days
    )

<# TODO
    1) Add exceptions section    
#####
#>
    
    $style = "<style>BODY{font-family: Arial; font-size: 10pt;}"
    $style = $style + "TABLE{border: 1px solid black; border-collapse: collapse;}"
    $style = $style + "TH{border: 1px solid black; background: #dddddd; padding: 5px; }"
    $style = $style + "TD{border: 1px solid black; padding: 5px; }"
    $style = $style + "</style>"

$ADSplat = @{
    Filter = "*"
    Searchbase = $Searchbase
    ResultPageSize = 2000 
    resultSetSize = $null
    Properties = "*"
    }

$ADServers = Get-ADComputer @ADSplat | where {$_.Enabled -eq "True"} | Select Name

$Table = @()

$ADServers.Name | % {
    $Machine = $_
    $Result = Test-Connection -ComputerName $Machine -Count 1 -Quiet
   
    If ($Result -like "False"){
        $Obj = [PSCustomObject]@{
           Name = $Machine
           Result = "Offline"
        }
        $Table += $Obj
    }

}

$Table | ConvertTo-Csv | Out-file c:\Temp\OfflineAD.csv
$Attach = Get-Item C:\Temp\OfflineAD.csv

Send-MailMessage -To "alex.glasbey@paconsulting.com" -From "infadm201@paconsulting.com" -Subject "AD Computers - Offline" -SmtpServer "SMTPSERVER" `
-Body ($Table | ConvertTo-Html -Head $style | Out-String ) -BodyAsHtml -Attachments $Attach
}