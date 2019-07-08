asnp Citrix*

$style = "<style>BODY{font-family: Arial; font-size: 10pt;}"
$style = $style + "TABLE{border: 1px solid black; border-collapse: collapse;}"
$style = $style + "TH{border: 1px solid black; background: #D24B4B; padding: 5px; }"
$style = $style + "TD{border: 1px solid black; padding: 5px; }"
$style = $style + "</style>"

(Get-BrokerCatalog) | %{
    $CatalogName = $_.Name
    $Domain = $Env:UserDomain

    $desktops = Get-BrokerDesktop -CatalogName $CatalogName | Select MachineName, IPAddress, LastConnectionUser, LastConnectionTime, AssociatedUserFullNames, AssociatedUserNames, OSType, SummaryState

    $Table = @()

    $Desktops | % {
       $A = $_.AssociatedUserNames -Join " ; " | Out-String
       $B = $_.AssociatedUserFullNames -Join " ; " | Out-String

       $obj = [PSCustomObject]@{
       MachineName = $_.MachineName
       IPAddress = $_.IPAddress
       LastConnectionUser = $_.LastConnectionUser
       LastConnectionTime = $_.LastConnectionTime
       AssociatedUserFullNames = $B
       AssociatedUserNames = $A
       OSType = $_.OSType
       SummaryState = $_.SummaryState
        }
        $Table += $Obj
    } 
        
        $Table | Convertto-HTML -Head $style | Out-File ("c:\temp\" + ($Domain + "_" + $CatalogName + "-Audit" + ".htm"))
}