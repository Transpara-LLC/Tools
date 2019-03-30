#Imports
Import-Module WebAdministration
Import-Module SQLPS -DisableNameChecking
[system.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
#end imports

#global variables
$global:siteName=""
$global:instanceName=""
$global:newInstanceName=""
$global:pathToSQLDATA=""
$global:newSQLDBName=""



#----------------------dont edit variables below--------------




$global:pathToInstance=""
$global:pathToOldInstance=""
$global:newCacheServiceName=""
$global:newAlertServiceName=""
$global:SQLDBName=""
$global:SQLServer=""
$global:SQLTable = "dbo.tableInterfaces"
$global:remoteContextPath = "" 
$global:remoteContextExists = ""
$networkServicePassword = (new-object System.Security.SecureString)
$global:NetworkServiceCredentials = New-Object System.Management.Automation.PSCredential ("NT AUTHORITY\NETWORK SERVICE", $networkServicePassword)
#end of global variables


function getStartupType($serviceName) {
    $startType = Get-WmiObject -Class Win32_Service -Property StartMode -Filter "Name='$serviceName'" | select -property * -ExcludeProperty PSComputerName, __GENUS, __CLASS, __SUPERCLASS, __DYNASTY, __RELPATH, __PROPERTY_COUNT, __DERIVATION, __SERVER, __NAMESPACE, __PATH, PSComputerName, Scope, Path, Options, ClassPath, Properties, SystemProperties, Qualifiers, Site, Container
    $startType = $startType -replace "StartMode", ""
    $startType = $startType -replace "@{=", ""
    $startType = $startType -replace "}", ""
    if($startType -eq "Manual") {
        return "Manual"
    }else {
        return "Automatic"
    }
    return "Automatic"
}
function getServiceStatus($serviceName) {
   return (Get-Service -Name $serviceName).Status 
}
function hasInputs {
    if($global:siteName -ne "" -and $global:instanceName -ne "" -and $global:newInstanceName -ne "" -and $global:newSQLDBName -ne "" ) {
         return 0
    }
    return 1    
}
function getInputs { 
    if($global:siteName -eq "") {
        $global:siteName = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the name of the site where the VKPI instance is", "Site Name Request")
    }
    if($global:instanceName -eq "") {   
         $global:instanceName =  [Microsoft.VisualBasic.Interaction]::InputBox("Enter the old instance name", "Instance Name Request")
    }
    if($global:newInstanceName -eq "") {
        $global:newInstanceName = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the new instance name", "Instance Name Request")
    }
    $global:newSQLDBName = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the new SQL DB name (leave blank for same as VKPI instance name) ", "SQL DB Name Request")
    if($global:newSQLDBName -eq "") {
        $global:newSQLDBName =  $global:newInstanceName 
    }
    if($global:pathToSQLDATA -eq ""){
     $global:pathToSQLDATA=[Microsoft.VisualBasic.Interaction]::InputBox("Enter the new SQL Data folder Example: C:\Program Files\Microsoft SQL Server\MSSQL12.SQLEXPRESS\MSSQL\DATA ", "SQL Data Folder Path Request")
    }
}
function createBaseServices {
    $global:newCacheServiceName="Visual KPI Cache Server ($global:newInstanceName)"
    $global:newAlertServiceName="Visual KPI Alert Server ($global:newInstanceName)"
    $oldCacheName  ="Visual KPI Cache Server ($global:instanceName)"
    $oldAlertName = "Visual KPI Alert Server ($global:instanceName)"
    $cacheStartType = getStartupType($oldCacheName)
    $alertStartType = getStartupType($oldAlertName)
    New-Service -Name "$global:newCacheServiceName" -BinaryPathName "$global:pathToNewInstance\Cache Server\VisualKPICache.exe" -DisplayName "$global:newCacheServiceName" -StartupType $cacheStartType -Description "Enables caching of KPI data in Visual KPI Server" -Credential $global:NetworkServiceCredentials
    New-Service -Name "$global:newAlertServiceName" -BinaryPathName "$global:pathToNewInstance\Alerter\VisualKPIAlerter.exe" -DisplayName "$global:newAlertServiceName" -StartupType $alertStartType -Description "Enables sending of Alerts based on KPI state changes in Visual KPI Server" -Credential $global:NetworkServiceCredentials
    $cacheStatus=getServiceStatus($oldCacheName)
    $alertStatus=getServiceStatus($oldAlertName)
    if($cacheStatus -eq "Running") {
        Start-Service -Name $newCacheServiceName
    }
    if($alertStatus -eq "Running") {
        Start-Service -Name $newAlertServiceName
    } 
}

function cloneDatabase {

    Backup-SqlDatabase -ServerInstance "$global:SQLServer" -Database "$global:SQLDBName"
    [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.SMO') | out-null
    $s = New-Object ('Microsoft.SqlServer.Management.Smo.Server') "$global:SQLServer" 
    $backupPath= $s.Settings.BackupDirectory
    
    $DBBackupPath = "$backupPath\$SQLDBName.bak"



    $database = "$global:newSQLDBName"
    $backupLocation =  $DBBackupPath
   
    $dataFileLocation ="$global:pathToSQLDATA\$database.mdf"
    $logFileLocation = "$global:pathToSQLDATA\$database.ldf"
    $log="_log"
$sql = @"
select user_name(),suser_sname()
GO  
USE [master]


RESTORE DATABASE [$database] 
FROM DISK = N'$backupLocation' 
WITH FILE = 1,  
     MOVE N'$global:SQLDBName' TO N'$dataFileLocation',  
     MOVE N'$global:SQLDBName$log' TO N'$logFileLocation',  
     NOUNLOAD, REPLACE, STATS = 5

ALTER DATABASE [$database] 
    SET MULTI_USER
"@

    invoke-sqlcmd $sql -ServerInstance $global:SQLServer 
}
function createRemoteContextServices {
    if(!$global:remoteContextExists){
        "No RCS"
        return
    }
    $RemoteContextFolders = Get-ChildItem "$global:pathToNewInstance\Remote Context Services"
    foreach($folder in $RemoteContextFolders) {
        "$folder"
        $FolderChild = get-childItem "$global:pathToNewInstance\Remote Context Services\$folder" 
        foreach($child in $FolderChild) {
            $childName = Get-ItemProperty -path "$global:pathToNewInstance\Remote Context Services\$folder\$child"  | Select-Object Name
            $childName = $childName -replace "@{Name=", ""
            $childName = $childName -replace "}", ""
            if($childName.endswith("Service.exe") ) {
                $folderName = Get-ItemProperty -path "$global:pathToNewInstance\Remote Context Services\$folder"  | Select-Object Name
                $folderName = $folderName -replace "@{Name=", ""
                $folderName = $folderName -replace "}", ""
                $serviceExecutableName = $childName
                $remoteContextServiceName = ""
                $remoteContextServiceBinaryPath= ""
                $remoteContextServiceDescription =""
                if($serviceExecutableName.contains("SQL")) {
                    $remoteContextServiceName = "Visual KPI SQL RC Server ($global:newInstanceName : $folderName)"
                    $remoteContextServiceDescription ="Enables automatic creation of Visual KPI Components from a source SQL database"
                } elseif($serviceExecutableName.contains("NIWNAS")) {
                    $remoteContextServiceName = "Visual KPI NIWNAS RC Server ($global:newInstanceName : $folderName)"
                    $remoteContextServiceDescription ="Enables automatic creation of Visual KPI Components from a NIWNAS server"
                } elseif($serviceExecutableName.contains("Monday")) {
                    $remoteContextServiceName = "Visual KPI Monday RC Server ($global:newInstanceName : $folderName)"
                    $remoteContextServiceDescription ="Enables automatic creation of Visual KPI Components from a Monday.com site"
                } elseif($serviceExecutableName.contains("inmation")) {
                    $remoteContextServiceName = "Visual KPI inmation RC Server ($global:newInstanceName : $folderName)"
                    $remoteContextServiceDescription ="Enables automatic creation of Visual KPI Components from inmation Models"
                } elseif($serviceExecutableName.contains("AFIntegration")) {
                    $remoteContextServiceName = "Visual KPI AF RC Server ($global:newInstanceName : $folderName)"
                    $remoteContextServiceDescription ="Enables automatic creation of Visual KPI Components from AF Elements"
                } 
                $remoteContextServiceBinaryPath= "$global:pathToNewInstance\Remote Context Services\$folder\$child"
                $oldServiceName = $remoteContextServiceName -replace "$global:newInstanceName", "$global:instanceName"   
                $startType = getStartupType($oldServiceName)
                $oldServiceStatus = getServiceStatus($oldServiceName)
                New-Service -Name "$remoteContextServiceName" -BinaryPathName "$remoteContextServiceBinaryPath" -DisplayName "$remoteContextServiceName" -StartupType $startType -Description "$remoteContextServiceDescription" -Credential $global:NetworkServiceCredentials
                if($oldServiceStatus -eq "Running") {
                     Start-Service -Name $remoteContextServiceName
    
                }
              
            }
        }
    }
}
function cloneAppPoolsAndFiles {
    $ws="_WS"
    copy "IIS:\AppPools\TAP_$global:instanceName" "IIS:\AppPools\TAP_$global:newInstanceName"
    copy "IIS:\AppPools\TAP_$global:instanceName$ws" "iis:\apppools\TAP_$global:newInstanceName$ws"
    $physicalPathToNewInstance = $global:pathToInstance -replace "$global:instanceName", "$global:newInstanceName"
    $correctSitePath = $global:pathToInstance -replace "$global:instanceName", ""
    Set-ItemProperty "IIS:\Sites\$global:siteName\"  physicalPath "$correctSitePath"


    copy $global:pathToInstance -Destination $physicalPathToNewInstance -Recurse -Force

    ConvertTo-WebApplication "IIS:\Sites\$global:siteName\$global:newInstanceName" -ApplicationPool "TAP_$global:newInstanceName"
    "appool set"
    ConvertTo-WebApplication "IIS:\Sites\$global:siteName\$global:newInstanceName\Interfaces" -ApplicationPool "TAP_$global:newInstanceName"
  
    ConvertTo-WebApplication "IIS:\Sites\$global:siteName\$global:newInstanceName\Webservice" -ApplicationPool "TAP_$global:newInstanceName$ws"
    $interfaces = Get-ChildItem "iis:\Sites\$global:siteName\$global:instanceName\Interfaces"
    foreach($interface in $interfaces) {
        $interfaceName = Get-ItemProperty -path "$interface"  | Select-Object Name
        $interfaceName = $interfaceName -replace "@{Name=", ""
        $interfaceName = $interfaceName -replace "}", ""
        $appPoolName = Get-ItemProperty -path "iis:\Sites\$siteName\$instanceName\Interfaces\$interfaceName" | Select-Object applicationPool 
        $appPoolName = $AppPoolName -replace "@{applicationPool=", ""
        $appPoolName = $AppPoolName -replace "}", ""
        $newAppPoolName = $appPoolName -replace "$global:instanceName","$global:newInstanceName"
        copy "IIS:\AppPools\$appPoolName" "IIS:\AppPools\$newAppPoolName" 
        ConvertTo-WebApplication "IIS:\Sites\$global:siteName\$global:newInstanceName\Interfaces\$interfaceName" -ApplicationPool $newAppPoolName 
    }
}
function stopAppPools {
    $appPools = Get-ChildItem "iis:\AppPools"
    foreach($appPool in $appPools) {
        $appPoolName = $appPool | Select-Object Name
        $appPoolName = $AppPoolName -replace "@{name=", ""
        $appPoolName = $AppPoolName -replace "}", ""  
        if($appPoolName.contains("$global:instanceName")) {
            echo $appPoolName
            Stop-WebAppPool $appPoolName
        }
    }
}
function startAppPools {
    $appPools = Get-ChildItem "iis:\AppPools"
    foreach($appPool in $appPools) {
        $appPoolName = $appPool | Select-Object Name
        $appPoolName = $AppPoolName -replace "@{name=", ""
        $appPoolName = $AppPoolName -replace "}", ""  
        echo $appPoolName
        Start-WebAppPool $appPoolName
    }
}

function replaceConnectionStrings {
    (Get-Content -Path "$global:pathtoNewInstance\web.config") | ForEach-Object {$_ -replace ";DataBase=$global:SQLDBName;",";DataBase=$global:newSQLDBName;"} | Set-Content -Path "$global:pathtoNewInstance\web.config"
    (Get-Content -Path "$global:pathtoNewInstance\Cache Server\VisualKPICache.xml") | ForEach-Object {$_ -replace ";DataBase=$global:SQLDBName;",";DataBase=$global:newSQLDBName;"} | Set-Content -Path "$global:pathtoNewInstance\Cache Server\VisualKPICache.xml"
    (Get-Content -Path "$global:pathtoNewInstance\Alerter\VisualKPIAlerter.xml") | ForEach-Object {$_ -replace ";DataBase=$global:SQLDBName;",";DataBase=$global:newSQLDBName;"} | Set-Content -Path "$global:pathtoNewInstance\Alerter\VisualKPIAlerter.xml"
    
    if(!$global:remoteContextExists){
        "No RCS"
        return
    }
    $remotecontextFolders = Get-ChildItem "$global:pathToNewInstance\Remote Context Services"
    foreach($folder in $RemoteContextFolders) {
        "$folder"
         $FolderChild = get-childItem "$global:pathToNewInstance\Remote Context Services\$folder" 
         foreach($child in $FolderChild) {
            $childName = Get-ItemProperty -path "$global:pathToNewInstance\Remote Context Services\$folder\$child"  | Select-Object Name
            $childName = $childName -replace "@{Name=", ""
            $childName = $childName -replace "}", ""
            if($childName.endswith("Service.xml") ) {
             (Get-Content -Path "$global:pathtoNewInstance\Remote Context Services\$folder\$child") | ForEach-Object {$_ -replace ";DataBase=$global:SQLDBName;",";DataBase=$global:newSQLDBName;"} | Set-Content -Path "$global:pathtoNewInstance\Remote Context Services\$folder\$child"
            }
         }
    }
}

function editDatabaseInterfacesURL { 
    $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
    $SqlConnection.ConnectionString = "Server = $global:SQLServer; Database = 
    $global:newSQLDBName; Integrated Security = True"
    $SqlConnection.Open()
    $update = @"
     UPDATE $global:SQLTable
     SET 
     [URL] = REPLACE([URL], '$global:instanceName/interfaces/', '$global:newInstanceName/interfaces/') 
"@
    $dbwrite = $SqlConnection.CreateCommand()
    $dbwrite.CommandText = $update
    $dbwrite.ExecuteNonQuery() 
    $Sqlconnection.Close()
    "Edited Interface SQL Table"
}
function getServerAndDB {
  $lines = Get-Content -path "$global:pathToInstance\web.config"
  $serverPattern = "Server=(.*);DataBase"
  $dbPattern = "DataBase=(.*);Integrated"
  foreach($line in $lines) {
    if($line.contains("connectionString=")) {
      $server = [regex]::Match($line, $serverPattern)
      $server = $server.value
      $server = $server -replace "Server=", ""
      $server = $server -replace ";DataBase", ""
      $global:SQLServer= $server
      $db = [regex]::Match($line, $dbPattern)
      $db = $db.value
      $db = $db -replace "DataBase=", ""
      $db = $db -replace ";Integrated", ""
      $global:SQLDBName= $db
    }
  }
}
function mapDatabaseRoleToLogin {
    #get variables
    $instanceName = "$global:SQLServer"
    $loginName = "NT AUTHORITY\NETWORK SERVICE"
    $dbUserName = "NT AUTHORITY\NETWORK SERVICE"
    $databasename = $global:newSQLDBName
    $roleName = "db_owner"
    $server = New-Object -TypeName Microsoft.SqlServer.Management.Smo.Server -ArgumentList $instanceName
    $server.Logins.Refresh()
    #$server.Logins | Format-Table -Property Parent, ID, Name, CreateDate,  LoginType
    #add a database mapping
    $database = $server.Databases[$databasename]
    $login = $server.Logins[$loginName]
    if ($database.Users[$dbUserName])
    {
        $database.Users[$dbUserName].Drop()
    }
    $dbUser = New-Object `
    -TypeName Microsoft.SqlServer.Management.Smo.User `
    -ArgumentList $database, $loginName
    $dbUser.Login = $loginName
    $dbUser.Create()

    #assign database role for a new user
    $dbrole = $database.Roles[$roleName]
    $dbrole.AddMember($dbUserName)
    $dbrole.Alter
}
try {

getInputs
if(hasInputs -eq 1){return}
#cloneIISApp #clones the file system stuff
$global:pathToInstance = Get-WebFilePath "iis:\Sites\$global:siteName\$global:instanceName"
getServerAndDB
stopAppPools

$cango = ""
do
{
    $cango = "go"
    "Attempting to stop app pools"
    Start-Sleep -Seconds 1
    $appPools = Get-ChildItem "iis:\AppPools"
    foreach($appPool in $appPools) {
        $appPoolName = $appPool | Select-Object Name
        $appPoolName = $AppPoolName -replace "@{name=", ""
        $appPoolName = $AppPoolName -replace "}", ""  
        if($appPoolName.contains("$global:instanceName")) {
            if((Get-WebAppPoolState -Name $appPoolName).Value -eq "Running") {
                $cango = ""
            }
        }
    }
}
until ($cango -eq "go")
"App Pool  stopped"

cloneAppPoolsAndFiles

$global:pathToNewInstance = Get-WebFilePath "iis:\Sites\$global:siteName\$global:newInstanceName"
$global:remoteContextPath = "$global:pathToInstance\Remote Context Services" 
$global:remoteContextExists = Test-Path -LiteralPath $remoteContextPath


createRemoteContextServices
createBaseServices
replaceConnectionStrings
"We will now clone the database, this might take a while... please standby"
cloneDatabase
editDatabaseInterfacesURL
mapDatabaseRoleToLogin
startAppPools
} catch {"This is embarrasing, an error has ocurred"}