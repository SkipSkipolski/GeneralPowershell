$VMs = Get-VM #Add Server

[array]$VMInfo=@()

$style = "<style>BODY{font-family: Arial; font-size: 10pt;}"
$style = $style + "TABLE{border: 1px solid black; border-collapse: collapse;}"
$style = $style + "TH{border: 1px solid black; background: #dddddd; padding: 5px; }"
$style = $style + "TD{border: 1px solid black; padding: 5px; }"
$style = $style + "</style>"

$VMs | %  {
    $VM = $_.Name
    $annotation = $_ | Get-annotation

    $obj=New-Object PsObject

    Add-Member -InputObject $obj -MemberType NoteProperty -Name VMname -Value $VM
    Add-Member -InputObject $obj -MemberType NoteProperty -Name OTRS_Ref -Value $Annotation[0].Value
    Add-Member -InputObject $obj -MemberType NoteProperty -Name Role -Value $Annotation[1].Value
    Add-Member -InputObject $obj -MemberType NoteProperty -Name Department -Value $Annotation[2].Value
    Add-Member -InputObject $obj -MemberType NoteProperty -Name Contact -Value $Annotation[3].Value
    
    $VMInfo+=$obj
}


#Mail items
    $str = "Hi,

You are receiving this email as you/your team has been identified as the owner of one or more of the virtual machines in the attached HTML file.

Please can you do the following:

1)  Let me know which servers are still required. 
2)  If you believe the server belongs to another user/team let me know.
3)  Supply any additional information for the machine if you would like it included in the details to help identify the machine in future.

The process for identifying VM ownership is still in development so apologies in advance if you have been assigned something erroneously.

If you need any more information please let me know.

Many thanks"

    $Grouping = $VMInfo | Group "Contact"

    $Grouping | ForEach-Object {
    If($_.Group.Contact -eq ""){
        
        $Data = $_.Group | Select VMName, "OTRS_Ref", "Role", "Department", "Contact", "VMRequired" | ConvertTo-Html -Head $Style | Out-File "C:\Temp\ErrVM.htm"
        $HTML = Get-Item "C:\Temp\ErrVM.htm"

        Send-MailMessage -To "admin@company.com" -From "ITDepartment@company.com" -Subject "REQ: VM Ownership - ERROR_NoVMName" -SmtpServer "SMTPSERVER" `
        -Body "Attached VMs Missing Owner Annotation - Please Investigate" -Attachments $HTML 
        #Break
    }
    Else{
        
        $usrSearch = @()
            
        $UsrSearch += ($_.Name).Split("/").TrimStart(" ").TrimEnd(" ")

        $mailAddr = $UsrSearch | % { Get-ADUser -Filter {Name -like $_ } -Properties Mail | Select Mail}

        $Data = $_.Group | Select VMName, "OTRS_Ref", "Role", "Department", "Contact", "VMRequired"

        $Data | ConvertTo-Html -Head $style | Out-File ("C:\Temp\" + $usrSearch + "_" + "VMOwnership" + ".htm")
        $HTML = Get-Item ("C:\Temp\" + $usrSearch + "_" + "VMOwnership" + ".htm")    

        If($mailAddr.Mail){
            Send-MailMessage -To "admin@company.com" -From "ITDepartment@company.com" -Subject "VM Ownership - $(($mailAddr.Mail).Replace("@company.com",";"))" -SmtpServer "SMTPSERVER" `
            -Body $Str -Attachments $HTML
        }else{
            Send-MailMessage -To "admin@company.com" -From "ITDepartment@company.com" -Subject "VM Ownership - ERROR" -SmtpServer "SMTPSERVER" `
            -Body $Str -Attachments $HTML
        }
    }
}

<#Region AD User Check
$Grouping.Name | % {
    $UsrSearch += ($_).Split("/").TrimStart(" ").TrimEnd(" ")
}
#>

    $Grouping | ForEach-Object {
    If($_.Group.Contact -eq ""){
        
        $Data = $_.Group | Select VMName, "OTRS_Ref", "Role", "Department", "Contact", "VMRequired" | ConvertTo-Html -Head $Style | Out-File "C:\Temp\ErrVM.htm"
        $HTML = Get-Item "C:\Temp\ErrVM.htm"

        Send-MailMessage -To "admin@company.com" -From "ITDepartment@company.com" -Subject "VM Ownership - ERROR_NoVMName" -SmtpServer "SMTPSERVER" `
        -Body "Attached VMs Missing Owner Annotation - Please Investigate" -Attachments $HTML 
        #Break
    }
    Else{
        
        $usrSearch = @()
            
        $UsrSearch += ($_.Name).Split("/").TrimStart(" ").TrimEnd(" ")

        $mailAddr = $UsrSearch | % { Get-ADUser -Filter {Name -like $_ } -Properties Mail | Select Mail}

        $Data = $_.Group | Select VMName, "OTRS_Ref", "Role", "Department", "Contact", "VMRequired"

        $Data | ConvertTo-Html -Head $style | Out-File ("C:\Temp\" + $usrSearch + "_" + "VMOwnership" + ".htm")
        $HTML = Get-Item ("C:\Temp\" + $usrSearch + "_" + "VMOwnership" + ".htm")    

        If($mailAddr.Mail){
            Send-MailMessage -To "admin@company.com" -From "ITDepartment@company.com" -Subject "REQ: VM Ownership - $(($mailAddr.Mail).Replace("@PACONSULTING.COM",";"))" -SmtpServer "SMTPSERVER" `
            -Body $Str -Attachments $HTML
        }else{
            Send-MailMessage -To "admin@company.com" -From "ITDepartment@company.com" -Subject "REQ: VM Ownership - ERROR" -SmtpServer "SMTPSERVER" `
            -Body $Str -Attachments $HTML
        }
    }
}
