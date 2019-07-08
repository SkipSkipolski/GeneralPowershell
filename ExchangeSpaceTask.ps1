#Whitespace report to be run from Exchange shell



#All space reports can be run from CAM-BLA-211 (except Edge servers)


$CASServer = "CAMCAS201",
"CAMCAS202",
"CAMCAS203",
"CAMCAS204",
"CAMCAS205",
"CAMCAS206"

$BlaServer = "CAM-BLA-211",
"CAM-BLA-313"

$MailboxServer = "CAMMBX101",
"CAMMBX102",
"LONMBX101"

#region Space check
[array]$Table = @()

$MailboxServer | Foreach {
    $wmi = get-wmiobject "Win32_LogicalDisk" -ComputerName $_ #| Where {$_.DeviceID -like "C:" -or $_.DeviceID -eq "D:"}
$ServerName = $_

    $WMI | % {
        $Obj = [PSCustomObject]@{
            Server = $ServerName
            Drive = $_.DeviceID
            DriveName = $_.VolumeName
            FreeSpace = $_.FreeSpace/1GB
        }
    $Table += $Obj
    }

}
#endregion

#region WhiteSpace
$Databases = $BlaServer + $MailboxServer

$databases | % { Get-MailboxDatabase -Server $_ -Status | select Server, Name, AvailableNewMailboxSpace}


