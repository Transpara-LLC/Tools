#PREREQES : THE SQLPS Module comes from a SQL Server installation, if there is none you can try install-module sqlserver and then import-module sqlserver (You need to have powershell 5 for installing)

#Imports
Import-Module WebAdministration
Import-Module SQLPS -DisableNameChecking
[system.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
#end imports

#global hardcodable variables
$global:siteName="Default Web Site"
$global:sourceInstanceName=""
$global:newInstanceName=""
$global:newSQLDBName=""
#end of global variables
















#----------------------dont edit variables below--------------


$networkServicePassword = (new-object System.Security.SecureString)
$global:NetworkServiceCredentials = New-Object System.Management.Automation.PSCredential ("NT AUTHORITY\NETWORK SERVICE", $networkServicePassword)
$global:SQLTable = "dbo.tableInterfaces" #the table to edit the interfaces url
$global:pathToInstance=""#the path to the source instance
$global:sourceSQLDBName=""#will be obtained and set by looking on the file system webconfig
$global:SQLServer=""#will be obtained and set by looking on the file system webconfig
$global:remoteContextPath = ""#the path of the remote context services 
$global:remoteContextExists = ""# a variable to check if there are remote context services


#gets the startup type of a service
function getServiceStartupType($serviceName) {
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

#gets the status of a service (running or stopped)
function getServiceStatus($serviceName) {
   return (Get-Service -Name $serviceName).Status 
}

#returns 1 if an input required to execute the program is missing. It is used to stop the execution
function hasInputs {
    if($global:siteName -ne "" -and $global:sourceInstanceName -ne "" -and $global:newInstanceName -ne "" -and $global:newSQLDBName -ne ""  ) {
         return 0
    }
    return 1    
}

#gets all the user inputs if they are not already hardcoded, if the new SQL database name is not entered, it will be the same as the new instance name
function getInputs { 
    if($global:siteName -eq "") {
        $global:siteName = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the name of the site where the VKPI instance is", "Site Name Request")
    }
    if($global:sourceInstanceName -eq "") {   
         $global:sourceInstanceName =  [Microsoft.VisualBasic.Interaction]::InputBox("Enter the old instance name", "Instance Name Request")
    }
    if($global:newInstanceName -eq "") {
        $global:newInstanceName = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the new instance name", "Instance Name Request")
    }
    if($global:newSQLDBName -eq "") {
        $global:newSQLDBName = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the new SQL DB name (leave blank for same as VKPI instance name) ", "SQL DB Name Request")
        if($global:newSQLDBName -eq "") {
            $global:newSQLDBName =  $global:newInstanceName 
        }
    }

}

function createBaseServices {
    $newCacheServiceName="Visual KPI Cache Server ($global:newInstanceName)"
    $newAlertServiceName="Visual KPI Alert Server ($global:newInstanceName)"
    $oldCacheName  ="Visual KPI Cache Server ($global:sourceInstanceName)"
    $oldAlertName = "Visual KPI Alert Server ($global:sourceInstanceName)"
    $cacheStartType = getServiceStartupType($oldCacheName)
    $alertStartType = getServiceStartupType($oldAlertName)
    New-Service -Name "$newCacheServiceName" -BinaryPathName "$global:pathToNewInstance\Cache Server\VisualKPICache.exe" -DisplayName "$newCacheServiceName" -StartupType $cacheStartType -Description "Enables caching of KPI data in Visual KPI Server" -Credential $global:NetworkServiceCredentials
    New-Service -Name "$newAlertServiceName" -BinaryPathName "$global:pathToNewInstance\Alerter\VisualKPIAlerter.exe" -DisplayName "$newAlertServiceName" -StartupType $alertStartType -Description "Enables sending of Alerts based on KPI state changes in Visual KPI Server" -Credential $global:NetworkServiceCredentials
    $cacheStatus=getServiceStatus($oldCacheName)
    $alertStatus=getServiceStatus($oldAlertName)
    if($cacheStatus -eq "Running") {
        Start-Service -Name $newCacheServiceName
    }
    if($alertStatus -eq "Running") {
        Start-Service -Name $newAlertServiceName
    } 
}

#makes a backup of the source database and restores it for the new instance
function cloneDatabase {
    $log="_log"
    Backup-SqlDatabase -ServerInstance "$global:SQLServer" -Database "$global:sourceSQLDBName" -Initialize
    [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.SMO') | out-null
    $s = New-Object ('Microsoft.SqlServer.Management.Smo.Server') "$global:SQLServer" 
    $backupDirectoryPath= $s.Settings.BackupDirectory
    $MDFDirectory = $s.Settings.DefaultFile
    $LDFDirectory = $s.Settings.DefaultLog

    $DBBackupPath = "$backupDirectoryPath\$sourceSQLDBName.bak"
    $dataFileLocation ="$MDFDirectory\$global:newSQLDBName.mdf"
    $logFileLocation = "$LDFDirectory\$global:newSQLDBName$log.ldf"
    $sql = @"
    select user_name(),suser_sname()
    GO  
    USE [master]
    RESTORE DATABASE [$global:newSQLDBName] 
    FROM DISK = N'$DBBackupPath' 
    WITH FILE = 1,  
        MOVE N'$global:sourceSQLDBName' TO N'$dataFileLocation',  
        MOVE N'$global:sourceSQLDBName$log' TO N'$logFileLocation',  
        NOUNLOAD, REPLACE, STATS = 5
    GO
    ALTER DATABASE [$global:newSQLDBName] MODIFY FILE ( NAME = '$global:sourceSQLDBName', NEWNAME = '$global:newSQLDBName');
    ALTER DATABASE [$global:newSQLDBName] MODIFY FILE ( NAME = '$global:sourceSQLDBName$log', NEWNAME = '$global:newSQLDBName$log');
    ALTER DATABASE [$global:newSQLDBName] SET MULTI_USER 
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
                $oldServiceName = $remoteContextServiceName -replace "$global:newInstanceName", "$global:sourceInstanceName"   
                $startType = getServiceStartupType($oldServiceName)
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
    copy "IIS:\AppPools\TAP_$global:sourceInstanceName" "IIS:\AppPools\TAP_$global:newInstanceName"#copies the main app pool (TAP_INSTANCEName)

    copy "IIS:\AppPools\TAP_$global:sourceInstanceName$ws" "iis:\apppools\TAP_$global:newInstanceName$ws"#Copies the Web service app pool

    $physicalPathToNewInstance = $global:pathToInstance -replace "$global:sourceInstanceName", "$global:newInstanceName"

    $correctSitePath = $global:pathToInstance -replace "$global:sourceInstanceName", ""
    Set-ItemProperty "IIS:\Sites\$global:siteName\"  physicalPath "$correctSitePath" 

    copy $global:pathToInstance -Destination $physicalPathToNewInstance -Recurse -Force #copies the files 

    ConvertTo-WebApplication "IIS:\Sites\$global:siteName\$global:newInstanceName" -ApplicationPool "TAP_$global:newInstanceName" #converts the instance folder into a web app and assigns an app pool

    ConvertTo-WebApplication "IIS:\Sites\$global:siteName\$global:newInstanceName\Interfaces" -ApplicationPool "TAP_$global:newInstanceName"#converts the interfaces folder into a web app and assigns an app pool
  
    ConvertTo-WebApplication "IIS:\Sites\$global:siteName\$global:newInstanceName\Webservice" -ApplicationPool "TAP_$global:newInstanceName$ws"#converts the webservice folder into a web app and assigns an app pool
    $interfaces = Get-ChildItem "iis:\Sites\$global:siteName\$global:sourceInstanceName\Interfaces" #gets all the interfaces
    foreach($interface in $interfaces) {#copies all interfaces app pools and converts the folders into web applications
        $interfaceName = Get-ItemProperty -path "$interface"  | Select-Object Name
        $interfaceName = $interfaceName -replace "@{Name=", ""
        $interfaceName = $interfaceName -replace "}", ""
        $appPoolName = Get-ItemProperty -path "iis:\Sites\$siteName\$global:sourceInstanceName\Interfaces\$interfaceName" | Select-Object applicationPool 
        $appPoolName = $AppPoolName -replace "@{applicationPool=", ""
        $appPoolName = $AppPoolName -replace "}", ""
        $newAppPoolName = $appPoolName -replace "$global:sourceInstanceName","$global:newInstanceName"
        copy "IIS:\AppPools\$appPoolName" "IIS:\AppPools\$newAppPoolName" 
        ConvertTo-WebApplication "IIS:\Sites\$global:siteName\$global:newInstanceName\Interfaces\$interfaceName" -ApplicationPool $newAppPoolName 
    }
}

#stops the app pools of the source instance
function stopAppPools {
    $appPools = Get-ChildItem "iis:\AppPools"
    foreach($appPool in $appPools) {
        $appPoolName = $appPool | Select-Object Name
        $appPoolName = $AppPoolName -replace "@{name=", ""
        $appPoolName = $AppPoolName -replace "}", ""  
        if($appPoolName.contains("$global:sourceInstanceName")) {#stops the app pool that contain the source instance name in their name
            echo $appPoolName
            Stop-WebAppPool $appPoolName
        }
    }
}

#starts the app pools for the source and new instance
function startAppPools {
    $appPools = Get-ChildItem "iis:\AppPools"
    foreach($appPool in $appPools) {
        $appPoolName = $appPool | Select-Object Name
        $appPoolName = $AppPoolName -replace "@{name=", ""
        $appPoolName = $AppPoolName -replace "}", ""  
        if($appPoolName.contains("$global:sourceInstanceName") -or $appPoolName.contains("$global:newInstanceName")) {#starts the app pools that contain the source and new instance name in their name
            echo $appPoolName
            Start-WebAppPool $appPoolName
        }
    }
}

#replaces the connection strings of the cloned files (for the instance, rcs, cache server and alerter)
function replaceConnectionStrings {
    (Get-Content -Path "$global:pathtoNewInstance\web.config") | ForEach-Object {$_ -replace ";DataBase=$global:sourceSQLDBName;",";DataBase=$global:newSQLDBName;"} | Set-Content -Path "$global:pathtoNewInstance\web.config"
    (Get-Content -Path "$global:pathtoNewInstance\Cache Server\VisualKPICache.xml") | ForEach-Object {$_ -replace ";DataBase=$global:sourceSQLDBName;",";DataBase=$global:newSQLDBName;"} | Set-Content -Path "$global:pathtoNewInstance\Cache Server\VisualKPICache.xml"
    (Get-Content -Path "$global:pathtoNewInstance\Alerter\VisualKPIAlerter.xml") | ForEach-Object {$_ -replace ";DataBase=$global:sourceSQLDBName;",";DataBase=$global:newSQLDBName;"} | Set-Content -Path "$global:pathtoNewInstance\Alerter\VisualKPIAlerter.xml"
    
    if(!$global:remoteContextExists){ #if the remote context folder doesnt exist then we stop here
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
             (Get-Content -Path "$global:pathtoNewInstance\Remote Context Services\$folder\$child") | ForEach-Object {$_ -replace ";DataBase=$global:sourceSQLDBName;",";DataBase=$global:newSQLDBName;"} | Set-Content -Path "$global:pathtoNewInstance\Remote Context Services\$folder\$child"
            }
         }
    }
}

#updates interfaces.dbo table to point to the new instance 
function editDatabaseInterfacesURL { 
    $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
    $SqlConnection.ConnectionString = "Server = $global:SQLServer; Database = 
    $global:newSQLDBName; Integrated Security = True"
    $SqlConnection.Open()
    $update = @"
     UPDATE $global:SQLTable
     SET 
     [URL] = REPLACE([URL], '$global:sourceInstanceName/interfaces/', '$global:newInstanceName/interfaces/') 
"@
    $dbwrite = $SqlConnection.CreateCommand()
    $dbwrite.CommandText = $update
    $dbwrite.ExecuteNonQuery() 
    $Sqlconnection.Close()
    "Edited Interface SQL Table"
}

#searches for the SQL server and the database inside the web.config file
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
      $global:sourceSQLDBName= $db
    }
  }
}

#adds the login to the database so that it can be accessed 
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

getInputs #
if(hasInputs -eq 1){return}

$global:pathToInstance = Get-WebFilePath "iis:\Sites\$global:siteName\$global:sourceInstanceName"

"--------------------------------------Getting server and database--------------------------------------"
getServerAndDB

"--------------------------------------Stoping app pools--------------------------------------"
stopAppPools

#waits until the app pools are stopped
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
        if($appPoolName.contains("$global:sourceInstanceName")) {
            if((Get-WebAppPoolState -Name $appPoolName).Value -eq "Running") {
                $cango = ""
            }
        }
    }
}
until ($cango -eq "go")
"--------------------------------------App Pools stopped--------------------------------------"



"--------------------------------------Cloning file system--------------------------------------"
cloneAppPoolsAndFiles

$global:pathToNewInstance = Get-WebFilePath "iis:\Sites\$global:siteName\$global:newInstanceName"
$global:remoteContextPath = "$global:pathToInstance\Remote Context Services" 
$global:remoteContextExists = Test-Path -LiteralPath $remoteContextPath

"--------------------------------------Creating remote cs--------------------------------------"
createRemoteContextServices
"--------------------------------------Creating base services--------------------------------------"
createBaseServices
"--------------------------------------Replacing connection strings--------------------------------------"
replaceConnectionStrings
"--------------------------------------Cloning database--------------------------------------"
cloneDatabase
"--------------------------------------Editing interfaces.dbo--------------------------------------"
editDatabaseInterfacesURL
"Creating database login"
mapDatabaseRoleToLogin
"--------------------------------------Starting app pools--------------------------------------"
startAppPools
} catch {"This is embarrasing, an error has ocurred"}
