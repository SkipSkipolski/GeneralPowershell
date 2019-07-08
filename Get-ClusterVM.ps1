$VIServer = Connect-VIServer

Function Get-VMAnnotation{

    $AllVMs = (Get-VM -Server $VIServer)

    [array]$VMs=@()

    $AllVMs | %  {
            $VM = $_.Name
            $annotation = $_ | Get-annotation

            $obj=New-Object PsObject

            Add-Member -InputObject $obj -MemberType NoteProperty -Name VMname -Value $VM
            Add-Member -InputObject $obj -MemberType NoteProperty -Name OTRS_Ref -Value $Annotation[0].Value
            Add-Member -InputObject $obj -MemberType NoteProperty -Name Role -Value $Annotation[1].Value
            Add-Member -InputObject $obj -MemberType NoteProperty -Name Department -Value $Annotation[2].Value
            Add-Member -InputObject $obj -MemberType NoteProperty -Name Contact -Value $Annotation[3].Value

            $VMs+=$obj
    }
    $VMs
}