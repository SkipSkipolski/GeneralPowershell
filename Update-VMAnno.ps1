$CSV = Import-Csv C:\temp\VMAnno.csv

$CSV[4] | % {
    $Issue = ""
    Try { $VM = $_.VMName
    Get-VM $VM -EA "Stop" | Out-Null 
    Write-Host ("Virtual Machine currently {0}" -f $VM)
    } 
    Catch {
        Write-Host ("Error updating VM {0}" -f $VM) -ForegroundColor Red
        $Issue = 1
    }
    If (!$Issue){
    Set-Annotation -Entity $VM -CustomAttribute "1: OTRS Ref:"  -Value $_."1: OTRS Ref:"
    Set-Annotation -Entity $VM -CustomAttribute "2: Server Role:" -Value $_."2: Server Role:"
    Set-Annotation -Entity $VM -CustomAttribute "3: Contact Department:" -Value $_."3: Contact Department:"
    Set-Annotation -Entity $VM -CustomAttribute "4: Contact Names:" -Value $_."4: Contact Names:"
    }
}


#Get the Annotations for reference
$CSV[41..60] | % {Get-VM $_.VMName} | % {
    Get-Annotation -Entity $_ -CustomAttribute "1: OTRS Ref:" #-Value $_."1: OTRS Ref:";
    Get-Annotation -Entity $_ -CustomAttribute "2: Server Role:"# -Value $_."2: Server Role:";
    Get-Annotation -Entity $_ -CustomAttribute "3: Contact Department:" #-Value $_."3: Contact Department:";
    Get-Annotation -Entity $_ -CustomAttribute "4: Contact Names:"# -Value $_."4: Contact Names:";    
    }
    
    #TEST EDIT
    #Testing Branching
    #Ignore this stuff
