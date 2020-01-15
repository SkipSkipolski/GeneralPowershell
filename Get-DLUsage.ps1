#Script to get all DL's not used in a set time.

#Setting Static Varibles
$LogFile = C:\Temp\Get-PADLStats.ps1

function Get-DLUsage
{
    $ErrorActionPreference = "Stop"
    $group = Get-DistributionGroup
    $table = @()

    $Group | ForEach-Object{
        $Count = Get-MessageTrackingLog -Start ((Get-Date).AddHours(-24)) -Recipients ($_).PrimarySMTPAddress | where {$_.EventID -eq "Receive"}
        $Members = Get-DistributionGroup $_ | Get-DistributionGroupMember
        #Write-Output ("Distribution List $_ has received {0} messages in the last 24 hours`n" -f $Result.count)

        $Obj = [PSCustomObject]@{
            DLName = $_
            MemberCount = ($Members).count
            Received = ($Count).count
            LastUsage = $count | sort Timestamp | select -Last 1 | select -expand Timestamp
        }
        $Table += $obj
    }
    $Table
}

function Get-O365DLUsage
{
    $ErrorActionPreference = "Stop"
    #$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Authentication Basic -Credential $365Cred -AllowRedirection
    $group = Get-MSOLGroup | where {$_.GroupType -eq "DistributionList"}
    $table = @()

    Try{
        $Group | ForEach-Object{
            $GroupName = $_.EmailAddress
            $GroupID = $_.ObjectID.Guid
            $Count = Get-MessageTrace -Start ((Get-Date).AddHours(-24)) -EndDate (Get-Date) -RecipientAddress $GroupName
            #$Members = Get-MSOLGroupMember -GroupObjectId $GroupID
            #Write-Output ("Distribution List $_ has received {0} messages in the last 24 hours`n" -f $Result.count)

            $Obj = [PSCustomObject]@{
                DLName = $GroupName
                #MemberCount = ($Members).count
                Received = ($Count).count
                LastUsage = $count | sort Timestamp | select -Last 1
            }
            $Table += $obj
        }
    }
    Catch{
        Write-output $_.Exception.message 
    }
    $Table
}