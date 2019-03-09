Import-Module WebAdministration
[system.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
$interfaceName = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the name of the interface to be changed", "Existing Interface Name Request")
$newInterfaceName = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the new interface name", "New Interface Name Request")
$siteName = "Transpara"
#$siteName = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the name of the site where the VKPI instance is", "Site Name Request")
$instanceName = "VisualKPI"
#$instanceName = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the name of the VKPI instance where the interface is", "Instance Name Request")
$SQLServer = "localhost\SQLExpress"
#$SQLServer = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the SQL Server for example: 'localhost\SQLExpress' ", "Enter the SQL Server ")
$SQLDB = "$instanceName"





#---------------------------------------------------------------------------dont change anything below this-----------------------

$SQLTable = "dbo.tableInterfaces"



#getting apppool name
$appPoolName = Get-ItemProperty -path "iis:\Sites\$siteName\$instanceName\Interfaces\$interfaceName" | Select-Object applicationPool 
$appPoolName = $AppPoolName -replace "@{applicationPool=", ""
$appPoolName = $AppPoolName -replace "}", ""
#end


Stop-WebAppPool -Name $appPoolName 
$cacheServiceName = "Visual KPI Cache Server ($instanceName)"
$alertServiceName = "Visual KPI Alert Server ($instanceName)"
Stop-Service -Name "$alertServiceName"
Stop-Service -Name "$cacheServiceName"
$cacheService = Get-Service -Name "$cacheServiceName"
$alertService = Get-Service -Name "$alertServiceName"
$pathToSite = Get-WebFilePath "IIS:\Sites\$siteName"#physical path

do
{
    "Attempting to stop app pool and services"
    Start-Sleep -Seconds 1
}
until ( (Get-WebAppPoolState -Name $appPoolName).Value -eq "Stopped" -and $cacheService.Status -ne "Running" -and $alertService.Status -ne "Running" )
"App Pool and services stopped"



#renaming the physical directories and changing physical paths
try {
    #renaming physical folder

    Rename-Item "$pathToSite\$instanceName\Interfaces\$interfaceName" "$pathToSite\$instanceName\Interfaces\$newInterfaceName" -ErrorAction 'Stop' 
    "Interface folder renamed"

    #change the actual interface name
    $appCmd = "C:\windows\system32\inetsrv\appcmd.exe"
    & $appCmd set app "$siteName/$instanceName/Interfaces/$interfaceName" -path:"/$instanceName/Interfaces/$newInterfaceName"

    #change physical path
    Set-ItemProperty -Path "IIS:\Sites\$siteName\$instanceName\Interfaces\$newInterfaceName" -Name "physicalPath" -Value "$pathToSite\$instanceName\Interfaces\$newInterfaceName"
    "Changed physical path" 

    #SQL dbo.interfaces editing
    $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
    $SqlConnection.ConnectionString = "Server = $SQLServer; Database = 
    $SQLDB; Integrated Security = True"
    $SqlConnection.Open()
    $update = @"
     UPDATE $SQLTable
     SET 
     [URL] = REPLACE([URL], '/interfaces/$interfaceName', '/interfaces/$newInterfaceName'),
     [Name] = '$newInterfaceName'
     WHERE [Name]  = '$interfaceName'
"@
    echo "interface name : $interfaceName"
    $dbwrite = $SqlConnection.CreateCommand()
    $dbwrite.CommandText = $update
    $dbwrite.ExecuteNonQuery() 
    $Sqlconnection.Close()
    "Edited Interface SQL Table"

} catch {
    "An error occurred"
}
#end






#Re-start everything
Start-WebAppPool -Name $appPoolName 
Start-Service -Name $cacheServiceName
Start-Service -Name $alertServiceName
"Started app pool and services"


pause
