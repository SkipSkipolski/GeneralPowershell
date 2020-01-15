 #Audit Servers for verious information - running services, processes, backups.
 
 $filepath = "C:\Temp"
 $name = "Audit"

 
 $Role = Get-WindowsFeature | Where {$_."InstallState" -eq "Installed"} | Select Name, DisplayName, InstallState
 $Role | ConvertTo-html  -Body "<H2>Installed Roles</H2>" > "$filepath\$name.html"

 Get-Service | Where {$_.Status -eq "Running"} | Select DisplayName, Status | ConvertTo-html  -Body "<H2>Services</H2>" >> "$filepath\$name.html"
 Get-Process | Select Name | ConvertTo-html  -Body "<H2>Processes</H2>" >> "$filepath\$name.html"

 If ($Role.name -match "Windows-Server-Backup"){
      
 Get-WBBackupSet | ConvertTo-html  -Body "<H2>Backups</H2>" >> "$filepath\$name.html"
 
 }

 #GetPrinters
  If ($Role.name -match "Print-Services"){
    
 Get-Printer | select Name, DriverName, PortName, Shared, Published | ConvertTo-html  -Body "<H2>Printers</H2>" >> "$filepath\$name.html"

 Get-PrinterDriver | select Name, Manufacturer | ConvertTo-html  -Body "<H2>Printer Drivers</H2>" >> "$filepath\$name.html"

}

#Get shared folders
Get-WmiObject -Class win32_share | Select Name, Path | ConvertTo-html  -Body "<H2>Shared Folders</H2>" >> "$filepath\$name.html"

#Installed Programs
Get-WmiObject Win32_Product | Select @{Name="Application Name";Expression={$_.Caption}},Version,Vendor  | sort 'Application Name' | ConvertTo-html  -Body "<H2>Installed Applications</H2>" >> "$filepath\$name.html"

#Get Volume Licensing
cd C:\Windows\System32\
$lic = cscript slmgr.vbs /dli
$Lic = $Lic | % {$_ + "<Br>"}
ConvertTo-Html -Body ("<br>" + "<H2>Volume Licencing</H2>" + "<br>" + "<pre>" + ($lic) + "</pre>") >> "$filepath\$name.html"


 #Check IIS
   If ($Role.name -match "Web-server" ){
    
 cd C:\Windows\System32\inetsrv
 $iis = .\appcmd.exe list site
 $iis += "`n"
 $iis += "`n"
 $iis += .\appcmd.exe list app
 ConvertTo-Html -Body ("<Br>" + "<H2>IIS Sites</H2>" + "<br>" + "<pre>" + $iis + "</pre>") >> "$filepath\$name.html"
}  