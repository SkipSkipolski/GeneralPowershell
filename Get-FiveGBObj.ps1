#Get all servers with less than 5GB space as an object

Function Get-DiskSpace 
{
    $Servers = Get-ADComputer -SearchBase "OU=Servers,DC=company,DC=com" -filter *
    $PSCred = (Get-Credential)
    $ErrorActionPreference = "Stop"
    $logFile = "c:\Temp\fails.log"
    $table = @();
    #$format = "FT"

    Clear-Content $logFile 

    $Servers.name | ForEach{
        try 
        {
        Invoke-Command -computername $_ -Credential $PSCred -Scriptblock {
        
        $Disk = Get-WMIObject Win32_LogicalDisk -filter "DeviceID='C:'" | select DeviceID, @{n="Size";e={[math]::Round($_.Size/1GB,2)}}, `
        @{n="FreeSpace";e={[math]::Round($_.FreeSpace/1GB,2)}} 
        
        $OS = Get-CimInstance Win32_OperatingSystem -Property *
        
        If($Disk.FreeSpace -gt 5.00){
            
            $Object = [PSCustomObject]@{"Computer" = $OS.csname;
                        OperatingSystem = $OS.Caption;
                        Disk = $Disk.DeviceID;
                        DiskSize = $Disk.Size;
                        FreeSpace = $Disk.FreeSpace;
                        }

                    $table += $Object
                    }
                Write-Output $table
                }                     
                #$object = New-Object PSObject -Property $ValueHash
        }
        catch 
        {               
            Add-Content $logFile -Value $Error[0].errordetails
        } 
    }
}