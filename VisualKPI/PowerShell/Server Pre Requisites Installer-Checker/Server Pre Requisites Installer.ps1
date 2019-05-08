[system.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
function installPreRequisites {
    ""
    "Installing web server - IIS "
    Install-WindowsFeature -Name Web-Server -IncludeManagementTools
    ""
    "Instaling Basic Authentication"
    #Install-WindowsFeature -Name Web-Dyn-Compression
    Install-WindowsFeature -Name Web-Basic-Auth
    ""
    "Installing Windows Authentication"
    Install-WindowsFeature -Name Web-Windows-Auth
    ""
    "Installing .NET extensibility"
    Install-WindowsFeature -Name Web-Net-Ext45 
    ""
    "Installing ASP"
    Install-WindowsFeature -Name Web-Asp-Net45 
    #Install-WindowsFeature -Name Web-Mgmt-Service 
    ""

    "Installing .NET 4.7.2"
    installDotNet472
}
function installDotNet472{
$source = "https://download.microsoft.com/download/6/E/4/6E48E8AB-DC00-419E-9704-06DD46E5F81D/NDP472-KB4054530-x86-x64-AllOS-ENU.exe";
$destination = "$PSScriptRoot\dotnet47.exe"
#Check if the installer is in the folder. If installer exist, replace it
If ((Test-Path $destination) -eq $false) {
    New-Item -ItemType File -Path $destination -Force
} 
#install software
Invoke-WebRequest $source -OutFile $destination
Start-Process -Wait -FilePath "$destination" -ArgumentList "/S" -PassThru
}
function checkPreRequisites {
    $feature =Get-WindowsFeature *Web-Server*
    $installed = $feature.Installed
    "Web-Server (IIS) installed:      $installed"

    $feature =Get-WindowsFeature *Web-Basic-Auth*
    $installed = $feature.Installed
    "Basic Authentication installed:  $installed"

    $feature =Get-WindowsFeature *Web-Windows-Auth*
    $installed = $feature.Installed
    "Windows Authentication:          $installed"

    $feature =Get-WindowsFeature *Web-Net-Ext45*
    $installed = $feature.Installed
    ".NET extensibility installed:    $installed"

    $feature =Get-WindowsFeature *Web-Asp-Net45*
    $installed = $feature.Installed
    "ASP installed:                   $installed"


    $installed = (Get-ItemProperty "HKLM:SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full").Release -ge 461814
    ".NET 4.7.2 installed:            $installed"
}

        


try {
#$action = [Microsoft.VisualBasic.Interaction]::InputBox("Enter install for installing or check for checking the installations", "Action Request")
$action = "install"
if($action -eq "check") {
    ""
    "-------CHECKING PREREQUISITES---------"
    checkPreRequisites
    "-------END CHECKING PREREQUISITES---------"
    ""

}
if($action -eq "install") {
    ""
    "-------INSTALLING PREREQUISITES-------"
    installPreRequisites
    "-------END INSTALLING PREREQUISITES-------"
    ""
}
pause
}catch {}
