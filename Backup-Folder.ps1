function Backup-Location {

    param (
        [Parameter(Mandatory=$True,
        HelpMessage="Specify a Folder, or enter AllUserItems to backup typical folders"
        )]
        $Source,

        [Parameter(
            HelpMessage="Destination for backup. If left empty it will default to C:\Users\User\Backup, and create a dated subfolder"
        )]
        $Destination
    )
    
    If (($Source)[-1] -eq "\"){
        $Source = $Source.TrimEnd("\")
    }


    If(!$Destination){
        $Destination = $env:userprofile + "\" + ($Source).Split("\")[-1] + "_" + (Get-Date -F ddMMyy_HHmm)
        New-Item -ItemType "Directory" -Path $Destination
    }

    If ($Source -eq "AllUserItems"){
        $Source = "$env:UserProfile\Desktop",
        "$env:UserProfile\Documents",
        "$env:UserProfile\Downloads",
        "$env:UserProfile\Favorites",
        "$env:UserProfile\Pictures",
        "C:\Scripts"
    }

    Try{
        Copy-Item -Path $Source -Destination $Destination -Force -Recurse -ErrorAction "Stop"
    }
    Catch{
        Write-Host ("Failed to backup {0}, See error below:" -f $Source) -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor "Red"
    }
   
}
