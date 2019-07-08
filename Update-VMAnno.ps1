$CSV = Import-Csv C:\temp\VMAnno.csv

function Set-VMAnno{

$CSV| % {
    $VM = $_.VMName
    Set-Annotation -Entity $VM -CustomAttribute "1: OTRS Ref" -Value $_."OTRS_Ref"
    Set-Annotation -Entity $VM -CustomAttribute "2: Server Role" -Value $_."Role"
    Set-Annotation -Entity $VM -CustomAttribute "3: Contact Department" -Value $_."Department"
    Set-Annotation -Entity $VM -CustomAttribute "4: Contact Names" -Value $_."Contact"
    }
}

function Get-VMAnno{

    $CSV | % {
        $VM = $_.VMName
        Get-Annotation -Entity $VM -CustomAttribute "1: OTRS Ref"
        Get-Annotation -Entity $VM -CustomAttribute "2: Server Role"
        Get-Annotation -Entity $VM -CustomAttribute "3: Contact Department"
        Get-Annotation -Entity $VM -CustomAttribute "4: Contact Names"
        }
}

#!!!! In dev/test, all custom attributes have a colon at the end of the name. E.g. "1: OTRS Ref:"
$CSV[0] | % {
    $VM = $_
    Get-Annotation -Entity $VM -CustomAttribute "4: Contact Names" #-Value $_."Sanjeev Sharma"
}
#!!!! In Prod they do not have a colon.

#Run command: $CSV = (Import-Csv C:\temp\VMAnno.csv)[0..2] 

#Region Final Check Command

$AnnoCheck = (Get-VM) | % {Get-Annotation -Entity $_ -CustomAttribute "2: Server Role", "3: Contact Department", "4: Contact Names"}
$AnnoCheck | Where {$_.Value -eq ""}
    #Results will show any empty annotations except for OTRS Ref

#endRegion
