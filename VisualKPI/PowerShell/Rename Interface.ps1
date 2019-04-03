function Use-RunAs 
{    
    # Check if script is running as Adminstrator and if not use RunAs 
    # Use Check Switch to check if admin 
     
    param([Switch]$Check) 
     
    $IsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()` 
        ).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator") 
         
    if ($Check) { return $IsAdmin }     
 
    if ($MyInvocation.ScriptName -ne "") 
    {  
        if (-not $IsAdmin)  
        {  
            try 
            {  
                $arg = "-file `"$($MyInvocation.ScriptName)`"" 
                Start-Process "$psHome\powershell.exe" -Verb Runas -ArgumentList $arg -ErrorAction 'stop'  
            } 
            catch 
            { 
                Write-Warning "Error - Failed to restart script with runas"  
                break               
            } 
            exit # Quit this session of powershell 
        }  
    }  
    else  
    {  
        Write-Warning "Error - Script must be saved as a .ps1 file first"  
        break  
    }  
} 
 
Use-RunAs  
 


Import-Module WebAdministration
[system.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
$siteName = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the name of the site where the VKPI instance is", "Site Name Request")
$instanceName = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the name of the VKPI instance where the interface is", "Instance Name Request")
$interfaceName = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the name of the interface to be changed", "Existing Interface Name Request")
$newInterfaceName = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the new interface name", "New Interface Name Request")
#$siteName = "Default Web Site"
#$instanceName = ""






















#---------------------------------------------------------------------------dont change anything below this-----------------------

$pathToInstance = Get-WebFilePath "iis:\Sites\$siteName\$instanceName"
#gets sqldb and server from web.config
  $lines = Get-Content -path "$pathToInstance\web.config"
  $serverPattern = "Server=(.*);DataBase"
  $dbPattern = "DataBase=(.*);Integrated"
  foreach($line in $lines) {
    if($line.contains("connectionString=")) {
      $server = [regex]::Match($line, $serverPattern)
      $server = $server.value
      $server = $server -replace "Server=", ""
      $server = $server -replace ";DataBase", ""
      $SQLServer= $server
      $db = [regex]::Match($line, $dbPattern)
      $db = $db.value
      $db = $db -replace "DataBase=", ""
      $db = $db -replace ";Integrated", ""
      $SQLDB= $db
     }
  }
#end


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
    $pathToInterface = Get-WebFilePath "iis:\Sites\$siteName\$instanceName\Interfaces\$interfaceName"
     "reach"
     "$pathToInterface"
    
    $pathToNewInterface =  $pathToInterface -replace "Interfaces\\$interfaceName", "Interfaces\$newInterfaceName"
    "reach"
    echo "$pathToNewInterface"
    Rename-Item "$pathToInterface" "$pathToNewInterface" -ErrorAction 'Stop' 
    "Interface folder renamed"
  


    #change the actual interface name
    $appCmd = "C:\windows\system32\inetsrv\appcmd.exe"
    & $appCmd set app "$siteName/$instanceName/Interfaces/$interfaceName" -path:"/$instanceName/Interfaces/$newInterfaceName"#virtual path
    #change physical path
    Set-ItemProperty -Path "IIS:\Sites\$siteName\$instanceName\Interfaces\$newInterfaceName" -Name "physicalPath" -Value "$pathToNewInterface"#first part is virtual path
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
#Get-WebFilePath iis:\Sites\Transpara\VisualKPI\Interfaces\ODBC

pause
