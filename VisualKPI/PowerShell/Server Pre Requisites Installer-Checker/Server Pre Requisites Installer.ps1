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
}
function checkPreRequisites {
    $feature =Get-WindowsFeature *Web-Server*
    $installed = $feature.Installed
    echo "Web-Server (IIS) installed:      $installed"

    $feature =Get-WindowsFeature *Web-Basic-Auth*
    $installed = $feature.Installed
    echo "Basic Authentication installed:  $installed"

    $feature =Get-WindowsFeature *Web-Windows-Auth*
    $installed = $feature.Installed
    echo "Windows Authentication:          $installed"

    $feature =Get-WindowsFeature *Web-Net-Ext45*
    $installed = $feature.Installed
    echo ".NET extensibility installed:    $installed"

    $feature =Get-WindowsFeature *Web-Asp-Net45*
    $installed = $feature.Installed
    echo "ASP installed:                   $installed"
}

        


try {
$action = [Microsoft.VisualBasic.Interaction]::InputBox("Enter install for installing or check for checking the installations", "Action Request")
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
