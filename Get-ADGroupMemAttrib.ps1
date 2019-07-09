Function Get-ADGroupMemAttrib {

    Param(
        [Parameter(Position=0,
        Mandatory=$True,
        ValueFromPipeline=$True,
        HelpMessage="Enter a valid AD group name")]
        [array]$GroupName,

        [Parameter(Position=1,
        Mandatory=$False,
        ValueFromPipeline=$False,
        HelpMessage= "Enter a folder name, or the default will be C:\Temp\")]
        $Folder = "C:\Temp\",
        
        [Parameter(Position=2,
        Mandatory=$False,
        ValueFromPipeline=$False,
        HelpMessage= "Enter a file prefix, or the default will be ADattrib_")]
        $File = "ADattrib_"
    )

    [array]$Table = @()

    $GroupName | Foreach-Object { 
        
        Try{
            $Group = Get-ADGroup $_ -ErrorAction "Stop"
            $mem = $Group | Get-ADGroupMember -ErrorAction "Stop"
        }
        Catch{ 
            Write-Host "Error Getting AD Group Members" -ForegroundColor "Red"
            Write-Host -ForegroundColor "Red" $_.Exception.Message
            Break
        }

        $mem.distinguishedName | % {
            $Obj = New-Object PSObject

            Try{
                $ADobj = Get-ADObject -Identity $_ -Properties samAccountName, extensionAttribute5 | Select samAccountName, extensionAttribute5 -ErrorAction "Stop"
            }
            Catch{
                Write-Host "Error Getting AD User" -ForegroundColor "Red"
                Write-Host -ForegroundColor "Red" $_.Exception.Message
                break
            }

            Add-Member -InputObject $obj -MemberType NoteProperty -Name User -Value $ADobj.samAccountName
            Add-Member -InputObject $obj -MemberType NoteProperty -Name extAttrib5 -Value $ADObj.extensionAttribute5
            Add-Member -InputObject $obj -MemberType NoteProperty -Name Group -Value $Group.Name

            $Table += $Obj
        }
    }
    $Table | ConvertTo-CSV -NoTypeInformation | Out-File ($Folder + $File + "_" + $Group.Name + ".CSV")
}