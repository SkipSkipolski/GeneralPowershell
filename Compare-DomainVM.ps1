#Script takes VM's from a hypervisor and compares if there is an equivelant VM in each domain if the VM's follow a naming convention.
#E.g. Confirms if there is a CAMEXC001, a TSTEXC001 and a DEVEXC001.

$style = "<style>BODY{font-family: Arial; font-size: 10pt;}"
$style = $style + "TABLE{border: 1px solid black; border-collapse: collapse;}"
$style = $style + "TH{border: 1px solid black; background: #00AEFF; padding: 5px; }"
$style = $style + "TD{border: 1px solid black; padding: 5px; }"
$style = $style + "</style>"

$VM = Get-VM
$DomGroup = $VM | Group-Object -Property {$_.Name.SubString(0,3)} | Where {$_.Name -eq "TST" -or $_.Name -eq "DEV" -or $_.Name -eq "CAM"}

$DomGroup.Name[0]| % {
    $gn = $_

    $Ref = $DomGroup | Where {$_.Name -eq $gn}
    $Dif = $DomGroup | where {$_.Name -ne $gn}

    $RefSubStr = (($Ref).Group.Name).Substring(3,3) | Select -Unique
    
    $Collection = @()

        $Dif | % {
            $DifDom = $_.Name
            $DifSubStr = (($_.Group.Name).Substring(3,3) | Select -Unique)
            
            $Comp = Compare-Object -ReferenceObject $RefSubStr -DifferenceObject $DifSubStr | Where {$_.SideIndicator -eq "<="}

            #Write-Host ("Domain {0} has the following VM's not present in other domains:" -f $gn)
            #Build VM search string and list all VM's

            $HostStrBuild = $Comp.InputObject | % {$gn + $_ + "*"}

            $Table = @()
            
            $HostStrBuild | % {
                $Searcher = $_
                $SubStr = $Searcher.Substring(3,3)
                $MissingVM = Get-VM | where {$_.Name -like $Searcher} | select -ExpandProperty Name
                $Domain = $gn

                $Obj = [PSCustomObject]@{
                    Domain = $Domain
                    VMName = ($MissingVM | Out-String)
                    ComparisonDomain = $difDom
                    SubString = $SubStr
                }

                $Table += $Obj    
            }

            $Collection += $Table
            #$str = ("<p>The following servers are in $($Ref.name) domain, but not $($difDom) : <br></p>")
            #$table | ConvertTo-Html -Body $str -head $style | Out-file ("C:\Temp\" + "$($Ref.name)" + "_To_" +  "$($difDom)" + ".htm")


        }
        $Collection
    # New-Variable -Name "Table$gn" -Value $Table -Force
}
