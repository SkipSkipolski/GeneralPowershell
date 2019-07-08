Function Get-DiskSpace{
    $Servers = Get-ADComputer -SearchBase "OU=Servers,DC=padev,DC=dev" -filter *
    $PSCred = (Get-Credential)
    $ErrorActionPreference = "Stop"
    $logFile = c:\Temp\fails.log
    $table = @();

    $Servers.name[0..10] | ForEach{
            try 
            {
                Invoke-Command -computername $_ -Credential $PSCred -Scriptblock
                {
                    #$table = @();
                    $Disk = Get-WMIObject Win32_LogicalDisk -filter "DeviceID='C:'" | select DeviceID, @{n="Size";e={[math]::Round($_.Size/1GB,2)}}, @{n="FreeSpace";e={[math]::Round($_.FreeSpace/1GB,2)}} 
                    $OS = Get-CimInstance Win32_OperatingSystem -Property *
                    If($Disk.FreeSpace -gt 5.00){
                    $ValueHash = @{ "Computer"=$OS.csname;
                                "Operating System"=$OS.Caption;
                                "Disk"=$Disk.DeviceID;
                                "DiskSize"=$Disk.Size;
                                "Free Space"=$Disk.FreeSpace;
                        }
                    }                    
                #$object = New-Object PSObject -Property $ValueHash
                Write-Output $valueHash
                }
            }
            catch 
            {
            Add-Content $Logfile -Value $Error[0].errordetails
            }
    }
}