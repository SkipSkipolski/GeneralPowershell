#Functions to get and set annotations on VMWare VM's from a CSV. Requires connection to hypervisor.

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