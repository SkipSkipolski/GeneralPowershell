$style = "<style>BODY{font-family: Arial; font-size: 10pt;}"
$style = $style + "TABLE{border: 1px solid black; border-collapse: collapse;}"
$style = $style + "TH{border: 1px solid black; background: #dddddd; padding: 5px; }"
$style = $style + "TD{border: 1px solid black; padding: 5px; }"
$style = $style + "</style>"

asnp "Citrix*"

$Domain = $Env:UserDomain
$fileLoc = "C:\Temp\" + $Domain
$fileChk = Get-Item ("C:\Temp\" + $Domain) -ErrorAction SilentlyContinue

If(!$fileChk){
New-Item -Type Directory -Path $fileLoc | out-Null
}


Function Get-CitrixAudit{

    $nonPoolCat = Get-BrokerCatalog | Where {$_.Name -notlike "*Pooled*"}

    ($nonPoolCat) | %{
        $CatalogName = $_.Name

        $desktops = Get-BrokerDesktop -CatalogName $CatalogName | Select MachineName, IPAddress, LastConnectionUser, LastConnectionTime, AssociatedUserFullNames, AssociatedUserNames, OSType, SummaryState

        $Table = @()

        $Desktops | % {
           $A = $_.AssociatedUserNames -Join " ; " | Out-String
           $B = $_.AssociatedUserFullNames -Join " ; " | Out-String

           If($_.LastConnectionTime -lt ((Get-Date).AddMonths(-3))){
                $_.LastConnectionTime = ($_.LastConnectionTime | Out-String)
                $_.LastConnectionTime = "#font"+$_.LastConnectionTime+"font#"
           }

           $obj = [PSCustomObject]@{
           MachineName = $_.MachineName
           IPAddress = $_.IPAddress
           LastConnectionUser = $_.LastConnectionUser
           LastConnectionTime = ($_ | % {$_.LastConnectionTime | Out-String })
           AssociatedUserFullNames = $B
           AssociatedUserNames = $A
           OSType = $_.OSType
           SummaryState = $_.SummaryState
            }
           $Table += $Obj
           $HTML = $Table | Convertto-HTML -Head $Style 
           $html = $html -replace "#font","<font color='red'>"
           $html = $html -replace "font#","</font>"
       
           $html | Out-File ("C:\Temp\" + $Domain + "\" + ($CatalogName + "-Audit" + ".htm"))

        }     
    }
}

Function Compress-Audit {

$Domain = $Env:UserDomain
$fileLoc = "C:\Temp\" + $Domain

$Source = $fileLoc
$destination = "C:\Temp" + "\" + $Domain + "_Citrix_Audit.zip"

Add-Type -AssemblyName "system.io.compression.filesystem"
[io.compression.zipfile]::CreateFromDirectory($Source, $Destination)

Remove-Item $source -Recurse -Force
}


Get-CitrixAudit
Compress-Audit
