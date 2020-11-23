#Requires -RunAsAdministrator
cls
Set-ExecutionPolicy RemoteSigned

$UpdateWorkPath = 'C:\Temp\upgrade\'
$SaveIISWebsite = 'C:\inetpub\wwwroot\BC130'
$BCServerInstallPath = 'C:\Program Files\Microsoft Dynamics 365 Business Central\130\Service'
$BCClientInstallPath = 'C:\Program Files (x86)\Microsoft Dynamics 365 Business Central\130\RoleTailored Client'
$GeneralInstallFolderName = "Microsoft Dynamics 365 Business Central"
$PathIISHelpServer = 'C:\inetpub\wwwroot\DynamicsNAV130Help'

$NavIde = "$BCClientInstallPath\finsql.exe"

$oldDB = 'Demo Database NAV (13-0)'
$newDB = 'Demo Database BC (14-0)' 
$SQLServer = 'DE-W10-BC13'
$SQLUser = "sa"
$SQLPassword = "rene"

$BCServerInstance = "BC130"

$DVDcv = "C:\Temp\BC130\Dynamics.365.BC.27233.DE.DVD"
$DVDnv = "C:\Temp\BC140\Dynamics.365.BC.44663.DE.DVD"

[bool] $prozess = $true 
do{
    cls
    Write-Host Upgrade assistant BC130 -> BC140
    Write-Host `n
    Write-Output -NoEnumerate "1 BC 140 install", "2 BC 130 uninstall", "3 BC 130 BackUp", "4 BC 130 export object", "5 BC 140 import License", "6 BC 140 uninstall", "0 abort"
    Write-Host `n
    $eingabe = Read-Host -Prompt "Please choose"
    switch($eingabe){
        default{Write-Host "Ungültige Eingabe"}        
        1{
            # BC 140 Installieren
            cls
            $eingabe = Read-Host -Prompt "Continue to install BC 140? (Y/N)"
            switch($eingabe){
                Y{
                    Write-host '----------------------------------------------------------'
                    Write-host '#1##### BC 140 install - START'
                    Write-host '----------------------------------------------------------'


                    #install NAV
                    Import-Module "$($DVDnv)\NavInstallationTools.psm1" -force -DisableNameChecking

                    Install-NAVComponent -ConfigFile "$UpdateWorkPath\config.xml" -Verbose
                    
                    
                    Export-NAVApplicationObject -DatabaseName  $oldDB -DatabaseServer "$($SQLServer)"  -Path "$($UpdateWorkPath)\Objects\all14-0tables.fob" -Filter 'Type=Table;' -LogPath "$($UpdateWorkPath)\Logs" -Confirm:$true -Force
                    Invoke-NAVDatabaseConversion -DatabaseServer "$($SQLServer)" -DatabaseName $newDB
                    Import-NAVApplicationObject -DatabaseName  $newDB -DatabaseServer "$($SQLServer)"  -Path "$($UpdateWorkPath)\Objects\all14-0tables.fob" -ImportAction Overwrite -SynchronizeSchemaChanges No -LogPath "$($UpdateWorkPath)\Logs" -Verbose


                    Set-NavServerConfiguration -ServerInstance $BCServerInstance -KeyName "DatabaseName" -KeyValue "Demo Database BC (14-0)"

                    Restart-NAVServerInstance $BCServerInstance


                    Sync-NAVTenant -ServerInstance $BCServerInstance -Mode Sync 

                    Write-host '----------------------------------------------------------'
                    Write-host '#1##### BC 140 install - ENDE'
                    Write-host '----------------------------------------------------------'
                    $eingabe = Read-Host -Prompt "Done enter any key"
                }
                N{Break;}
            }      
        }
        2{
            # BC 130 uninstall
            cls
            $eingabe = Read-Host -Prompt "Continue to uninstall BC 130? (Y/N)"
            switch($eingabe){
                Y{
                    # Import Module für BC Befehle
                    Import-Module "${env:ProgramFiles(x86)}\Microsoft Dynamics 365 Business Central\130\RoleTailored Client\NavModelTools.ps1" -force

                    Write-host '----------------------------------------------------------'
                    Write-host '#2##### BC 130 uninstall - START'
                    Write-host '----------------------------------------------------------'

                    #Get-NAVAppInfo $BCServerInstance | Uninstall-NAVApp -Verbose

                    #Get-NAVServerInstance | Set-NAVServerInstance -Stop -Verbose
                    #Get-NAVServerInstance | Remove-NAVServerInstance -Confirm:$false -Verbose

                    $uninstallcmd = "$($DVDcv)\setup.exe /quiet /uninstall /log $UpdateWorkPath\logs\uninstall.txt"
                    Invoke-Expression $uninstallcmd -Verbose
                    Remove-Item "C:\Program Files (x86)\$($GeneralInstallFolderName)" -Recurse -Verbose
                    Remove-Item "C:\Program Files\$($GeneralInstallFolderName)" -Recurse -Verbose
                    Remove-Item $PathIISHelpServer -Recurse -Verbose

                    Remove-Website -Name 'Microsoft Dynamics 365 Business Central Help'
                    Remove-Website -Name 'Microsoft Dynamics 365 Business Central Web Client'

                    Write-host '----------------------------------------------------------'
                    Write-host '#2##### BC 130 uninstall - ENDE'
                    Write-host '----------------------------------------------------------'
                    $eingabe = Read-Host -Prompt "Done enter any key"
                 }
                 N{Break;}
            }
        }
        3{
            # BC BackUp Start
            cls
            $eingabe = Read-Host -Prompt "Continue to BackUp BC 130? (Y/N)"
            switch($eingabe){
                Y{
                    Write-host '----------------------------------------------------------'
                    Write-host '#3##### BC BackUp - START'
                    Write-host '----------------------------------------------------------'

                    New-Item $UpdateWorkPath -Name "BackUp" -ItemType Directory -Force
                    New-Item "$UpdateWorkPath\BackUp" -Name "ConfigFiles" -ItemType Directory -Force
                    New-Item "$UpdateWorkPath\BackUp" -Name "Website" -ItemType Directory -Force

                    Copy-Item -Path "$($BCServerInstallPath)\CustomSettings.config" -Destination "$($UpdateWorkPath)\BackUp\ConfigFiles" -Force
                    Copy-item -Path "$($BCServerInstallPath)\Instances" -Destination "$($UpdateWorkPath)\BackUp\Instances" -Force -recurse
                    Copy-item -Path $SaveIISWebsite -Destination  "$($UpdateWorkPath)\BackUp\Website" -Force -recurse

                    Write-host '----------------------------------------------------------'
                    Write-host '#3##### BC BackUp - ENDE'
                    Write-host '----------------------------------------------------------'
                    $eingabe = Read-Host -Prompt "Done enter any key"
                 }
                 N{Break;}
            }
        }
        4{
            # BC 130 export object
            cls
             $eingabe = Read-Host -Prompt "Continue to export objects from BC 130? (Y/N)"
            switch($eingabe){
                Y{
                    # Import Module für BC Befehle
                    Import-Module "${env:ProgramFiles(x86)}\Microsoft Dynamics 365 Business Central\130\RoleTailored Client\NavModelTools.ps1" -force

                    Write-host '----------------------------------------------------------'
                    Write-host '#4##### BC 130 export object - START'
                    Write-host '----------------------------------------------------------'

                    Export-NAVApplicationObject -Path "C:\txt2al\export\test.txt" -DatabaseName "Demo Database NAV (13-0)" -DatabaseServer $SQLServer -Username $SQLUser -Password $SQLPassword -Force -LogPath "C:\txt2al\export\Log" -Filter 'Type=MenuSuite;Id=1010..1030' | Split-NAVApplicationObjectFile -Destination C:\txt2al\export\

                    Write-host '----------------------------------------------------------'
                    Write-host '#4#####  BC 130 export object - ENDE'
                    Write-host '----------------------------------------------------------'
                    $eingabe = Read-Host -Prompt "Done enter any key"
                 }
                 N{Break;}
            }
        }
        5{
            # BC 130 import License
            cls
            $eingabe = Read-Host -Prompt "Continue to import License? (Y/N)"
            switch($eingabe){
                Y{
                    # Import Module für BC Befehle
                    Import-Module "${env:ProgramFiles(x86)}\Microsoft Dynamics 365 Business Central\130\RoleTailored Client\NavModelTools.ps1" -force

                    Write-host '----------------------------------------------------------'
                    Write-host '#4##### BC 130 import License - START'
                    Write-host '----------------------------------------------------------'

                    Import-NAVServerLicense 'MicrosoftDynamicsNavServer$MyInstance' -LicenseData ([Byte[]]$(Get-Content -Path "$UpdateWorkPath\hoch.rein Storage GmbH DynBC14 2020-09-25.flf" -Encoding Byte)) -Database Master

                    Write-host '----------------------------------------------------------'
                    Write-host '#4#####  BC 130 import License - ENDE'
                    Write-host '----------------------------------------------------------'
                    $eingabe = Read-Host -Prompt "Done enter any key"
                 }
                 N{Break;}
            }
        }
        6{
            # BC 140 uninstall
            cls
            $eingabe = Read-Host -Prompt "Continue to uninstall BC 140? (Y/N)"
            switch($eingabe){
                Y{
                    # Import Module für BC Befehle
                    Import-Module "${env:ProgramFiles(x86)}\Microsoft Dynamics 365 Business Central\140\RoleTailored Client\NavModelTools.ps1" -force

                    Write-host '----------------------------------------------------------'
                    Write-host '#6##### BC 140 uninstall - START'
                    Write-host '----------------------------------------------------------'

                    #Get-NAVAppInfo "BC140" | Uninstall-NAVApp -Verbose

                    #Get-NAVServerInstance | Set-NAVServerInstance -Stop -Verbose
                    #Get-NAVServerInstance | Remove-NAVServerInstance -Confirm:$false -Verbose

                    $uninstallcmd = "$($DVDnv)\setup.exe /quiet /uninstall /log $UpdateWorkPath\logs\uninstall.txt"
                    Invoke-Expression $uninstallcmd -Verbose
                    Remove-Item "C:\Program Files (x86)\$($GeneralInstallFolderName)" -Recurse -Verbose
                    Remove-Item "C:\Program Files\$($GeneralInstallFolderName)" -Recurse -Verbose
                    Remove-Item $PathIISHelpServer -Recurse -Verbose

                    Remove-Website -Name 'Microsoft Dynamics 365 Business Central Help'
                    Remove-Website -Name 'Microsoft Dynamics 365 Business Central Web Client'

                    Write-host '----------------------------------------------------------'
                    Write-host '#6##### BC 140 uninstall - ENDE'
                    Write-host '----------------------------------------------------------'
                    $eingabe = Read-Host -Prompt "Done enter any key"
                 }
                 N{Break;}
            }
        }
        0{$prozess = $false}       
    }
}
while($prozess -eq "false")