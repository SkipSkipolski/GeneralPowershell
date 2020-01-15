#Whitespace report to be run from Exchange shell



#All space reports can be run from CAM-BLA-211 (except Edge servers)


$MBXServer = "TSTMBX201",
"TSTMBX202",
"TSTMBX203",
"TSTMBX204",
"TSTMBX205",
"TSTMBX206"

#region Space check
[array]$Table = @()

$MBXServer | Foreach {
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


