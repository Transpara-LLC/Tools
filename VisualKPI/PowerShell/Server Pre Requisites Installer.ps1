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

pause

