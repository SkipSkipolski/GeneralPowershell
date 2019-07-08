# Reset variables
$global:ibEndPoints = $null
$global:piaAppServers = $null
$global:piaWebServers = $null
$global:ibAppServers = $null
$global:ibWebServers = $null
$global:prcsSchedulers = $null
$global:checkInterval = 30 # seconds between checks for running processes
$global:autoServiceStatus = 'Automatic'
$global:disabledServiceStatus = 'Disabled'
$global:logFile = "$PSSCriptRoot\logs\PeopleSoftMaintenanceLog.log"
$global:pshrMasterScheduler = $null
$global:psfinMasterScheduler = $null
$global:psppmMasterScheduler = $null
$global:checkType = 'Check'
$global:shutDownType = 'Stop'
$global:startUpType = 'Start'

Function Main()
{
    Initialise 
    CheckServices

    $action = Read-Host("Enter 1 to start , 2 to shut down or any other key to exit")
    switch ($action)
        {
            1 { StartUp }
            2 { ShutDown }
            default 
                {
                    WriteHostAndLog "You chose to exit" $fgColorInfo
                    Exit
                }
        }
} #end function Main

Function Initialise
{
    # Make sure logs folder exists
    if (-Not (Test-Path -Path "$PSSCriptRoot\logs"))
    {
        New-Item -Path "$PSSCriptRoot\logs" -ItemType Directory
    }
    $dtString = (Get-Date).ToString('yyyyMMdd')
    $global:logFile = $global:logFile.Replace('.log', "$dtString.log")
    WriteEmptyLogLine
    WriteHostAndLog "Initialising script using log file $global:logFile" $fgColorInfo

    $global:maintenanceRequestType = $global:checkType
    # Import the file details
    $filePath = Get-FileName("$PSSCriptRoot")
    if ($filePath -eq "")
    {
        WriteHostAndLog "No File Chosen, exiting script" $fgColorError
        Exit
    }
    WriteHostAndLog "File Chosen $filePath" $fgColorInfo
    
    $logfileName = $filePath -replace '.*\\' # Remove the path
    $logfileName = $logfileName.Replace('.csv', "$dtString.log")
    $logfileName = "$PSSCriptRoot\logs\$logfileName"
    # $global:logFile =  "$PSSCriptRoot\logs\$fileName_

    WriteHostAndLog "SwitchingLogFile to $logfileName" $fgColorInfo
    $global:logFile = $logfileName
    WriteEmptyLogLine
    WriteHostAndLog "Starting script" $fgColorInfo
    
    $global:envDetails = Import-Csv $filePath

    # Create lists of each type
    $global:ibEndPoints = $global:envDetails | Where-Object {$_.Type -eq 'IBEndPoint'}
    $global:piaAppServers = $global:envDetails | Where-Object {$_.Type -eq 'PIAAppServer'}
    $global:piaWebServers = $global:envDetails | Where-Object {$_.Type -eq 'PIAWebServer'}
    $global:ibAppServers = $global:envDetails | Where-Object {$_.Type -eq 'IBAppServer'}
    $global:ibWebServers = $global:envDetails | Where-Object {$_.Type -eq 'IBWebServer'}
    $global:prcsSchedulers = $global:envDetails | Where-Object {$_.Type -eq 'ProcessScheduler'}
    $global:receiveLocations = $global:envDetails | Where-Object {$_.Type -eq 'BizTalkReceiveLocation'}
    ValidateInput
    WriteHostAndLog "" $fgColorInfo
}

Function Get-FileName($startFolder)
{   
    [System.Reflection.Assembly]::LoadWithPartialName(“System.windows.forms”) | Out-Null
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.initialDirectory = $startFolder
    $openFileDialog.filter = “CSV files (*.csv)| *.csv”
    $openFileDialog.ShowDialog() | Out-Null
    $openFileDialog.filename
} #end function Get-FileName

Function ValidateInput
{
    WriteHostAndLog 'Environment details' $fgColorInfo
    $columns = " ", "Type", "Environment", "Database", "ServerName", "ServiceName", "Domain", "StartType", "PrimaryMasterScheduler", "Input Validation"

    WriteInputColumns $columns $fgColorInfo
    foreach ($service in $global:envDetails)
    {
        $validation = 'Not validated'
        # Add validation if needed or rely on input file being correct
        $columns = " ", $service.Type, $service.Environment, $service.Database, $service.ServerName, $service.ServiceName, $service.Domain, $service.StartType, $service.PrimaryMasterScheduler, $validation
        WriteInputColumns $columns $fgColorInfo
    }
}

Function WriteInputColumns ($columns, $fgColor)
{
    $output = "{0,-3} {1,-25} {2,-12} {3,-10} {4,-10} {5,-85} {6,-12} {7,-12} {8,-25} {9,-14}" -f $columns
    <# Write-Host $output -ForegroundColor $fgColor
    $dttmString = (Get-Date).ToString('yyyyMMddThhmmss')
    $fileOutput = $dttmString + $output
    $fileOutput | Out-File -FilePath $global:logFile -Append
    #>
    WriteHostAndLog $output $fgColor

}

Function GetLoginCredentials
{
    $global:user = Read-Host "Enter user" -AsSecureString
    $global:password = Read-Host "Enter password:" -AsSecureString
}

Function Confirm()
{
    $confirm = Read-Host("Are you sure you want to continue? Enter y to continue or any other key to exit")
    if ($confirm.ToLower() -ne 'y')
        {
            WriteHostAndLog "You chose not to continue." $fgColorInfo
            Read-Host "Press any key to exit."
            Exit
        }
} #end function Confirm

# CheckProcesses
# SuspendProcessSchedulers
Function WebException
{
    try
    {
        $webException = $_.Exception
        $environment = $endPoint.Environment
        WriteHostAndLog "Error calling web service $soapAction for $environment at $url"  $fgColorError
        # Write-Host $exception -ForegroundColor $fgColorError
        # Try to cast to xml for IB reponse

        $errorXml = [xml]$webException.ErrorDetails.Message
        # IB Response will only exist if the request has contacted PeopleSoft IB
     #   try
     #   {
            $errorMessage = $errorXML.Envelope.Body.Fault.detail.IBResponse.DefaultMessage.'#cdata-section'
            if ($errorMessage -eq $Null)
            {
                Throw
            }
            WriteHostAndLog "`t$errorMessage" $fgColorError
    #    }
    #    catch 
    #    {
    #        # Ignore exception
    #    }
        $faultString = $errorXML.Envelope.Body.Fault.faultstring
        if ($faultString -ne 'null')
        {
            WriteHostAndLog "`t$faultString" $fgColorError
        }
    }
    catch # Response does not include IBResponse.DefaultMessage so probably not a PeopleSoft error
    {
        $errorMessage = $webException.Message
        if ($errorMessage -eq $null)
        {
            $errorMessage = $webException
        }
        WriteHostAndLog "`t$errorMessage" $fgColorError
    }

    switch ($soapAction)
    { 
        'SuspendProcessServers.V1' {WriteHostAndLog "Unable to Suspend Process Schedulers for $environment" $fgColorError }
        'CheckProcesses.V1' {WriteHostAndLog "Unable to Check Running Processes for $environment" $fgColorError }
    }
    Confirm
}

Function CheckProcesses
{
    $soapAction = "CheckProcesses.V1"
    $clearTextUsr = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($global:user))
    $clearTextPwd = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($global:password))
    $query = [xml] @"
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
    <soapenv:Header xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/">
    <wsse:Security soap:mustUnderstand="1" xmlns:soap="http://schemas.xmlsoap.org/wsdl/soap/" xmlns:wsse="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd">
        <wsse:UsernameToken wsu:Id="UsernameToken-1" xmlns:wsu="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd">
        <wsse:Username>$clearTextUsr</wsse:Username>
        <wsse:Password  Type="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-username-token-profile-1.0#PasswordText">$clearTextPwd</wsse:Password>
        </wsse:UsernameToken>
    </wsse:Security>
    </soapenv:Header>
    <soap:Body>
    <CheckProcesses__CompIntfc__PAC_PROCESSMONITOR />
    </soap:Body>
</soap:Envelope>
"@
    $clearTextUsr = $null
    $clearTextPwd = $null
    try
    {
        foreach ($endpoint in $global:ibEndPoints)
        {
          #  $endpoint = $global:ibEndPoints | Where-Object { $_.Environment -eq $environment}
            $url = $endpoint.ServiceName
            $environment = $endPoint.Environment
            if ($url -eq $null)
            {
                WriteHostAndLog "No url found to check running Process Monitor processes for $environment" $fgColorWarning
                Confirm
                Return
            }
        
            $processCount = 9999        

            while ($processCount -gt 0)
            {
            
                [xml] $result = Invoke-WebRequest $url  -Method Post -ContentType "text/xml" -Body $query -Headers @{SOAPAction= $soapAction}
                [xml] $processList = $result.Envelope.Body.CheckProcesses__CompIntfc__PAC_PROCESSMONITORResponse.'#text'
                $processCount = $processList.ProcessList.ChildNodes.Count
                WriteHostAndLog "Number of Processes running for $environment : $processCount" $fgColorInfo
                if ($processCount -ne 0)
                {
                    $date = Get-Date
                    #Write-Host 'Process' $process.ProcessName 'began processing at' $process.TimeProcessBegan
                    $processList.ProcessList.Process | Format-Table -AutoSize | Out-String|% {WriteHostAndLog "`t $_" $fgColorInfo}
                   # Write-Host $processList.ProcessList.Process | Format-Table
                
                 #   Write-Host " " $processList.ProcessList.Process
                    Write-Host ... -NoNewline
                    WriteHostAndLog "$date Checking every $checkInterval seconds" $fgColorInfo 
                    Start-Sleep -Seconds $checkInterval
                }
            }
        }
       # Write-Host " " $result.Envelope.Body.CheckProcesses__CompIntfc__PAC_PROCESSMONITORResponse.'#text' -Separator `t
       # Write-Host " " $result -Separator `t
        
    }
    catch
    {
        WebException
    }
    $query = $null
}

Function SuspendProcessSchedulers
{
    $soapAction = 'SuspendProcessServers.V1'
    $clearTextUsr = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($global:user))
    $clearTextPwd = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($global:password))
    $query = [xml] @"
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
    <soapenv:Header xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/">
    <wsse:Security soap:mustUnderstand="1" xmlns:soap="http://schemas.xmlsoap.org/wsdl/soap/" xmlns:wsse="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd">
        <wsse:UsernameToken wsu:Id="UsernameToken-1" xmlns:wsu="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd">
        <wsse:Username>$clearTextUsr</wsse:Username>
        <wsse:Password  Type="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-username-token-profile-1.0#PasswordText">$clearTextPwd</wsse:Password>
        </wsse:UsernameToken>
    </wsse:Security>
    </soapenv:Header>

    <soap:Body>
    <SuspendProcessServers__CompIntfc__PAC_PROCESSMONITOR />
    </soap:Body>
</soap:Envelope>
"@
    $clearTextUsr = $null
    $clearTextPwd = $null
    try
    {
        foreach ($endpoint in $global:ibEndPoints)
        {
          #  $endpoint = $global:ibEndPoints | Where-Object { $_.Environment -eq $environment}
            $url = $endpoint.ServiceName
            $environment = $endPoint.Environment
            if ($url -eq $null)
            {
                WriteHostAndLog "No url found to check suspend Process Schedulers for $environment" $fgColorWarning
                Confirm
                Return
            }
            $finished = 0
            While ($finished -eq 0)
            {
                $finished = 1
                [xml] $result = Invoke-WebRequest $url -Method Post -ContentType "text/xml" -Body $query -Headers @{SOAPAction= $soapAction}
                # $result.Envelope.Body.SuspendProcessServers__CompIntfc__PAC_PROCESSMONITORResponse
                [xml] $serverList = $result.Envelope.Body.SuspendProcessServers__CompIntfc__PAC_PROCESSMONITORResponse.'#text'
                $serverStatus = $serverList.ServerList.Server

            WriteHostAndLog "Process scheduler status for $environment" $fgColorInfo            
                # Write-Host "ServerName" "Status" "Master" -Separator `t
                foreach ($scheduler in $serverStatus)
                {
                    if ($scheduler.Status -eq 'Running')
                    {
                        $finished = 0
                    }
                    $columns = "", $scheduler.ServerName, $scheduler.Status, $scheduler.Master
                    WriteProcessSchedulerColumns $columns $fgColorInfo
                } 
            }
            # $serverStatus | Format-Table -AutoSize
       

            $masterScheduler = $serverStatus | Where-Object { $_.Master -eq 'Y'} 
            switch ($environment)
            {
                "PSHR" { $global:pshrMasterScheduler =  $masterScheduler.ServerName }
                "PSFIN" { $global:psfinMasterScheduler = $masterScheduler.ServerName }
                "PSPPM" { $global:psppmMasterScheduler = $masterScheduler.ServerName }
            }
        
        }
    }
    catch
    {
        WebException
    }
    $query = $null
}

Function WriteProcessSchedulerColumns ($columns, $fgColor)
{
    $output = "{0,-3} {1,-8} {2,-12} {3,-10}" -f $columns
    WriteHostAndLog $output $fgColor
}

Function CheckApplicationServerStartUp
{
    foreach ($service in $global:piaAppServers | Where-Object { $_.StartType -eq $global:autoServiceStatus } )
    {
        CheckApplicationServerProcesses $service
    }

    foreach ($service in $global:ibAppServers)
    {
        CheckApplicationServerProcesses $service
    }
}

Function CheckApplicationServerProcesses($service)
{
    $processName = 'PSAPPSRV.exe'
    $domainArg = '-D ' + $service.Domain + '\b' # Use the word boundary \b to specify the end so we need the word and, for example, PBF92DEV results do not include PBF92DEVIB

    $processes = Get-CimInstance Win32_Process -ComputerName $service.ServerName | Where-Object {$_.Name -match $processName -and $_.CommandLine -match $domainArg} 
    OutputProcessResults $processes $service $processName
}


function CheckProcessSchedulerStartUp
{
    foreach($service in $global:prcsSchedulers  | Where-Object { $_.StartType -eq $global:autoServiceStatus } )
    {
        CheckProcessSchedulerProcesses $service
    }
}

Function CheckProcessSchedulerProcesses($service)
{
    $processName = 'PSPRCSRV.exe'
    # Use the word boundary \b to specify the end so we need the word and, for example, PBF92DEV results do not include PBF92DEVIB
    $domainArg = '-PS ' + $service.Domain + '\b' 
    $databaseArg = '-CD ' + $service.Database + '\b' # Look for the -CD argument which holds the database name

    $processes = Get-CimInstance Win32_Process -ComputerName $service.ServerName | Where-Object {$_.Name -eq $processName -and $_.CommandLine -match $domainArg -and $_.CommandLine -match $databaseArg } 
    OutputProcessResults $processes $service $processName
}

Function OutputProcessResults($processes, $service, $processName)
{
    $serverName = $service.ServerName
    $domain = $service.Domain
    $database = $service.Database

    if ($processes -ne $null)
    {
        if ($global:maintenanceRequestType -eq $global:startUpType -or $global:maintenanceRequestType -eq $global:checkType )
        {
            $fgColor = $fgColorInfo
        }
        else
        {
            $fgColor = $fgColorError
        }

            
        if ($processes -is [array])
        {
            $processText = $processes.Length.ToString() + ' Processes Started' 
        }
        else
        {
            $processText = '1 Process Started' 
        }
        WriteHostAndLog  "`t $serverName $processName $domain $processText for database $database" $fgColor
    }
    else
    {
        if ($global:maintenanceRequestType -eq $global:shutDownType)
        {
            $fgColor = $fgColorInfo
        }
        else
        {
            $fgColor = $fgColorError
        }
        WriteHostAndLog "`t $serverName - No $processName processes started for $domain for database' $database" $fgColor
    }
}

function CheckServices
{
        
    # $services = $global:piaAppServers + $global:piaWebServers + $global:ibAppServers + $global:ibWebServers + $global:prcsSchedulers
    # $services
    WriteHostAndLog "Current Service status" $fgColorInfo
    WriteHostAndLog "Application servers, process schedulers and BizTalk receive locations will show as errors if you are restarting after a shut down" $fgColorInfo
    WriteServicesHeader $fgColorInfo
    # Check HR then FIN then PPM
    CheckServicesByEnvironment 'PSHR' $global:autoServiceStatus $fgColorInfo
    CheckServicesByEnvironment 'PSFIN'  $global:autoServiceStatus $fgColorInfo
    CheckServicesByEnvironment 'PSPPM' $global:autoServiceStatus $fgColorInfo

    WriteHostAndLog "Contingency Service status (not expected to be running):" $fgColorContingency
    WriteServicesHeader $fgColorContingency
    # Check HR then FIN then PPM
    CheckServicesByEnvironment 'PSHR' $global:disabledServiceStatus $fgColorContingency
    CheckServicesByEnvironment 'PSFIN' $global:disabledServiceStatus $fgColorContingency
    CheckServicesByEnvironment 'PSPPM' $global:disabledServiceStatus $fgColorContingency

    CheckBizTalkReceiveLocations
} #end function CheckServices

Function WriteServicesHeader($fgColor)
{
    WriteServiceStatusColumns "", "Type", "Server", "Service", "Status", "StartType", "Environment" $fgColor
}

Function CheckServicesByEnvironment($environment, $startType, $fgColorInfo)
{
     foreach ($service in $global:piaAppServers | Where-Object { $_.StartType -eq $startType -and $_.Environment -eq $environment } )
        {
            CheckService $service $fgColorInfo
        }
    foreach ($service in $global:ibAppServers | Where-Object { $_.StartType -eq $startType -and $_.Environment -eq $environment } )
        {
            CheckService $service $fgColorInfo
        }
    foreach ($service in $global:prcsSchedulers | Where-Object { $_.StartType -eq $startType -and $_.Environment -eq $environment } )
        {
            CheckService $service $fgColorInfo
        }
    foreach ($service in $global:piaWebServers | Where-Object { $_.StartType -eq $startType -and $_.Environment -eq $environment } )
        {
            CheckService $service $fgColorInfo
        }
    foreach ($service in $global:ibWebServers | Where-Object { $_.StartType -eq $startType -and $_.Environment -eq $environment } )
        {
            CheckService $service $fgColorInfo
        }

}

Function CheckService($service, $fgColor)
    {
        $svc = get-service -ComputerName $service.ServerName -Name $service.ServiceName -ErrorAction Ignore
        $serviceStatus = $svc.Status
        if ($service.Type -notmatch 'WebServer')
        {
            $fgColor = SetServiceStatusColor $serviceStatus $service
        }
        Else
        {
            $fgColor = $fgColorInfo
        }

        if ($svc -ne $null)
            {
                $startType= GetStartUpType $service.ServiceName
                WriteServiceStatusColumns "", $service.Type, $svc.MachineName, $svc.Name, $svc.Status, $startType, $service.Environment $fgColor
            }
        else
            {
                WriteServiceStatusColumns   "", $service.ServerName, $service.ServiceName, "Not Found", " ", " ", " " $fgColorError
            }
    } #end function CheckService

Function SetServiceStatusColor($serviceStatus, $service)
{
    if ($serviceStatus -eq 'Running' -or $serviceStatus -eq 'Enabled')
    {
        if ($service.StartType -eq 'Disabled')
        {
            Return $fgColorError
        }

        if ($global:maintenanceRequestType -eq $startUpType -or $global:maintenanceRequestType -eq $checkType)
        {
            Return $fgColorInfo
        }
        else
        {
            Return $fgColorError
        }
    }
    else
    {
        if ($service.StartType -eq 'Disabled')
        {
            Return $fgColorInfo
        }
        if ($global:maintenanceRequestType -eq $shutDownType)
        {
            Return $fgColorInfo
        }
        else
        {
            Return $fgColorError
        }
    }
}

Function GetStartUpType($serviceName)
    {
        Return (Get-CimInstance Win32_Service -ComputerName $service.ServerName -filter "Name='$serviceName'").StartMode
    } #end function GetStartUpType

Function WriteServiceStatusColumns ($columns, $fgColor)
    {
        $output = "{0,-3} {1,-25} {2,-10} {3,-60} {4,-10} {5,-10} {6,-12}" -f $columns
        <#
        Write-Host $output -ForegroundColor $fgColor
        $dttmString = (Get-Date).ToString('yyyyMMddThhmmss')
        $fileOutput = $dttmString + $output
        $fileOutput | Out-File -FilePath $global:logFIle
        #>
        WriteHostAndLog $output $fgColor
    } 

Function WriteHostAndLog($output, $fgColor)
{
    switch ($fgColor)
    {
        $fgColorError { $level = 'ERROR' }
        $fgColorWarning { $level = 'WARN ' }
        default { $level = 'INFO ' }
    }

    Write-Host $output -ForegroundColor $fgColor
    $dttmString = (Get-Date).ToString('yyyyMMdd HHmmss')
    $fileOutput = "$dttmString`t$level`t$output"
    $fileOutput | Out-File -FilePath $global:logFile -Append
}

Function WriteEmptyLogLine
{
    $fileOutput = ""
    $fileOutput | Out-File -FilePath $global:logFile -Append
}

Function ShutDown
{
    WriteHostAndLog "Are you sure you want to shut down?" $fgColorDefault
    Confirm
    $global:maintenanceRequestType = $shutDownType
    GetLoginCredentials
    SuspendProcessSchedulers
   # SuspendProcessSchedulers PSHR
    WriteHostAndLog "`tPSHR master scheduler: $global:pshrMasterScheduler" $fgColorInfo
    WriteHostAndLog "`tPSFIN master scheduler: $global:psfinMasterScheduler" $fgColorInfo
    WriteHostAndLog "`tPSPPM master scheduler: $global:psppmMasterScheduler" $fgColorInfo
    CheckProcesses
   # CheckProcesses PSHR
   # CheckServices
    WriteHostAndLog 'Checks complete starting shut down' $fgColorInfo
    DisableBizTalkReceiveLocations
    StopPSServices PSHR
    StopPSServices PSFIN
    WriteHostAndLog 'Checking application server and process scheduler processes have stopped - each entry should show no processes running' $fgColorInfo
    CheckApplicationServerStartUp
    CheckProcessSchedulerStartUp
    CheckServices
}

Function StopPSServices($environment)
{
    switch ($environment)
    {
        "PSHR" { $masterScheduler = $global:pshrMasterScheduler }
        "PSFIN" { $masterScheduler = $global:psfinMasterScheduler }
        "PSPPM" { $masterScheduler = $global:psppmMasterScheduler }
    }

    # Stop Process Schedulers
    $startType = $global:autoServiceStatus
    # Stop master scheduler last
    $services = $global:prcsSchedulers | Where-Object { $_.StartType -eq $startType -and $_.Environment -eq $environment -and $_.Domain -ne  $masterScheduler}
    WriteHostAndLog "Stopping slave process schedulers for $environment" $fgcolorInfo
    StopServices $services

    $services = $global:prcsSchedulers | Where-Object { $_.StartType -eq $startType -and $_.Environment -eq $environment -and $_.Domain -eq  $masterScheduler}
    WriteHostAndLog "Stopping master process scheduler for $environment" $fgcolorInfo
    StopServices $services

    $services = $global:piaAppServers | Where-Object { $_.StartType -eq $startType -and $_.Environment -eq $environment }
    WriteHostAndLog "Stopping PIA application servers for $environment" $fgcolorInfo
    StopServices $services

    $services = $global:ibAppServers | Where-Object { $_.StartType -eq $startType -and $_.Environment -eq $environment }
    WriteHostAndLog "Stopping IB application servers for $environment" $fgcolorInfo
    StopServices $services
}

Function StopServices($services)
{
    foreach ($service in $services) 
        { 
            $serviceName = $service.ServiceName
            $serverName = $service.ServerName
            $svc = get-service -ComputerName $serverName -Name $serviceName
            # Set service start type to Manual
            Set-Service -InputObject $svc -StartupType "Manual"
            WriteHostAndLog "`tStopping service $serviceName on $serverName" $fgColorInfo
            # Stop service
            Stop-Service -InputObject $svc -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
            <#
            While($svc.Status  -ne 'Stopped')
                { 
                    Write-Host 'Stopping...'-NoNewLine 
                    Start-Sleep 3 
                    $svc.Refresh()
                }
            $svc.Refresh() #>
            $startType = $svc.StartType
            $serviceStatus = $svc.Status
            WriteHostAndLog "`t`t$serverName`t$serviceName`t$serviceStatus`t$startType" $fgColorInfo
        }
}

Function Startup()
{
    $global:maintenanceRequestType = $startUpType
    WriteHostAndLog "Are you sure you want to start up?" $fgColorDefault
    Confirm
    StartPSServices PSHR
    StartPSServices PSFIN
    WriteHostAndLog "Waiting for 1 minute to allow processes to start"  $fgColorInfo
    Start-Sleep -Seconds 60
    WriteHostAndLog 'Checking application server and process scheduler processes have started up - each entry should show processes running' $fgColorInfo
    CheckApplicationServerStartUp
    CheckProcessSchedulerStartUp
    RestartWebServers
    EnableBizTalkReceiveLocations
    CheckServices
}

Function StartPSServices($environment)
{
    $startType = $global:autoServiceStatus
    $services = $global:ibAppServers | Where-Object { $_.StartType -eq $startType -and $_.Environment -eq $environment }
    StartServices $services
    $services = $global:piaAppServers | Where-Object { $_.StartType -eq $startType -and $_.Environment -eq $environment }
    StartServices $services
    $services = $global:prcsSchedulers | Where-Object { $_.StartType -eq $startType -and $_.Environment -eq $environment }
    StartServices $services
}

Function StartServices($services)
{
    foreach ($service in $services) 
    { 
        $serviceName = $service.ServiceName
        $serverName = $service.ServerName
        $svc = get-service -ComputerName $service.ServerName -Name $service.ServiceName
        # Set service start type to Automatic
        Set-Service -InputObject $svc -StartupType "Automatic"
        # Start service
        WriteHostAndLog "Starting service $serviceName on $serverName" $fgColorInfo
        Start-Service -InputObject $svc -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
        <#
            While($svc.Status  -ne 'Running')
            { 
                Write-Host 'Starting...'-NoNewLine 
                Start-Sleep 3 
                $svc.Refresh()
            }
        #>
        $svc.Refresh()
        
         # WriteHostAndLog "`t$svc.MachineName`t$svc.DisplayName`t$svc.Status`t$svc.StartType" -ForegroundColor $fgColorInfo
            $startType = $svc.StartType
            $serviceStatus = $svc.Status
            WriteHostAndLog "`t$serverName`t$serviceName`t$serviceStatus`t$startType" $fgColorInfo
    }
}

Function RestartWebServers
{
    foreach($service in $global:piaWebServers | Where-Object { $_.StartType -eq $global:autoServiceStatus })
    {
        $serviceName = $service.ServiceName
        $serverName = $service.ServerName
        $svc = get-service -ComputerName $serverName -Name $serviceName
        WriteHostAndLog "Restarting service $serviceName on $serverName" $fgColorInfo
        Restart-Service -InputObject $svc -WarningAction SilentlyContinue 
    }
}

Function CheckBizTalkReceiveLocations
{
    WriteHostAndLog "BizTalk status" $fgColorInfo
    WriteServiceStatusColumns "", "Type", "Server", "ReceiveLocation", "Status", "", "", "" $fgColorInfo
    foreach ($receiveLocation in $global:receiveLocations)
    {
        $recLocName = $receiveLocation.ServiceName
        $serverName = $receiveLocation.ServerName
       # WriteHostAndLog "Disabling $reclocName on $serverName" $fgColorInfo
       try
       {
           $recLoc = Get-CimInstance -ComputerName $serverName -ClassName MSBTS_ReceiveLocation -namespace 'root\MicrosoftBizTalkServer' -ErrorAction Stop | Where-Object { $_.Name -eq $recLocName } 
           if ( $recLoc.isDisabled -eq $true)
           {
                $status = 'Disabled'
                $fgColor = $fgColorError
           }
           Else
           {
                $status = 'Enabled'
                $fgColor = $fgColorInfo
           }
      
           $fgColor = SetServiceStatusColor $status $receiveLocation

           $columns = "", $receiveLocation.Type, $receiveLocation.ServerName, $receiveLocation.ServiceName, $status, "", "", ""
           WriteServiceStatusColumns $columns $fgColor
       }
       catch
       {
            WriteHostAndLog "`tCould not get receive location $recLocName on $serverName" $fgColorError
       }
    }
}

Function DisableBizTalkReceiveLocations
{
    WriteHostAndLog "Disabling BizTalk receive locations" $fgColorInfo
    foreach ($receiveLocation in $global:receiveLocations)
    {
        $reclocName = $receiveLocation.ServiceName
        $serverName = $receiveLocation.ServerName
        WriteHostAndLog "`tDisabling $reclocName on $serverName" $fgColorInfo
        try
        {
            $recLoc = Get-CimInstance -ComputerName $serverName -ClassName MSBTS_ReceiveLocation -namespace 'root\MicrosoftBizTalkServer' -ErrorAction Stop | Where-Object { $_.Name -eq $recLocName } 
            $result = Invoke-CimMethod -InputObject $recLoc -MethodName Disable                
        }
        catch
        {
            WriteHostAndLog "`tCould not disable receive location $recLocName on $serverName" $fgColorError
            Confirm
        }
    }
}

Function EnableBizTalkReceiveLocations
{
    WriteHostAndLog "Enabling BizTalk receive locations" $fgColorInfo
    foreach ($receiveLocation in $global:receiveLocations)
    {
        $reclocName = $receiveLocation.ServiceName
        $serverName = $receiveLocation.ServerName
        WriteHostAndLog "`tEnabling $reclocName on $serverName" $fgColorInfo
        try
        {
            $recLoc = Get-CimInstance -ComputerName $serverName -ClassName MSBTS_ReceiveLocation -namespace 'root\MicrosoftBizTalkServer' -ErrorAction Stop | Where-Object { $_.Name -eq $recLocName } 
            $result = Invoke-CimMethod -InputObject $recLoc -MethodName Enable                
        }
        catch
        {
            WriteHostAndLog "`tCould not enable receive location $recLocName on $serverName" $fgColorError
            Confirm
        }
    }
}

$fgColorDefault = "White"
$fgColorInfo = "Cyan"
$fgColorError = "Red"
$fgColorWarning = "Yellow"
$fgColorContingency = "Gray"

$fgColor = $fgColorDefault


Main


