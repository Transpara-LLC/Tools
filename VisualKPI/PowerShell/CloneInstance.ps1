
#TODO turn into applications the folders
#Todo create app pools and assign them (After cloning iis app)
#Todo edit interfaces.dbo to point to the new instance interface (After cloning DB) 
#todo point connection string to new DB
#Questions: where in the filesystem is the connection string of the server mgr
#answer: web.config in filesystem
#edit SQL interface neim

#Imports
Import-Module WebAdministration
Import-Module SQLPS -DisableNameChecking
[system.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
#end imports

#global variables
$global:siteName=""
$global:instanceName=""
$global:newInstanceName=""
$global:pathToInstance=""
$global:pathToOldInstance=""
$global:newCacheServiceName=""
$global:newAlertServiceName=""
$global:SQLDBName=""
$global:SQLServer=""
$global:SQLTable = "dbo.tableInterfaces"
$networkServicePassword = (new-object System.Security.SecureString)
$global:NetworkServiceCredentials = New-Object System.Management.Automation.PSCredential ("NT AUTHORITY\NETWORK SERVICE", $networkServicePassword)
#end of global variables

function hasInputs {
    if($global:siteName -ne "" -and $global:instanceName -ne "" -and $global:newInstanceName -ne "" -and $global:SQLDBName -ne "" -and $global:SQLServer -ne "") {
         return 0
    }
    return 1    
}
function getInputs { 
    $global:siteName = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the name of the site where the VKPI instance is", "Site Name Request")
    $global:instanceName =  [Microsoft.VisualBasic.Interaction]::InputBox("Enter the old instance name", "Instance Name Request")
    $global:newInstanceName = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the new instance name", "Instance Name Request")
    $global:SQLDBName = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the SQL DB name", "SQL DB Name Request")
    $global:SQLServer = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the SQL Server ", "SQL Server Request")
}
function createBaseServices {
    $global:newCacheServiceName="Visual KPI Cache Server ($global:newInstanceName)"
    $global:newAlertServiceName="Visual KPI Alert Server ($global:newInstanceName)"
    New-Service -Name "$global:newCacheServiceName" -BinaryPathName "$global:pathToNewInstance\Cache Server\VisualKPICache.exe" -DisplayName "$global:newCacheServiceName" -StartupType Automatic -Description "Enables caching of KPI data in Visual KPI Server" -Credential $global:NetworkServiceCredentials
    New-Service -Name "$global:newAlertServiceName" -BinaryPathName "$global:pathToNewInstance\Alerter\VisualKPIAlerter.exe" -DisplayName "$global:newAlertServiceName" -StartupType Automatic -Description "Enables sending of Alerts based on KPI state changes in Visual KPI Server" -Credential $global:NetworkServiceCredentials
    Start-Service -Name $newCacheServiceName
    Start-Service -Name $newAlertServiceName
}
function cloneDatabase {
    $SQLInstanceName = "$global:SQLServer"
    $SourceDBName   = "$global:SQLDBName"
    $CopyDBName = "$global:newInstanceName"
    $Server  = New-Object -TypeName 'Microsoft.SqlServer.Management.Smo.Server' -ArgumentList $SQLInstanceName
    $SourceDB = $Server.Databases[$SourceDBName]
    $CopyDB = New-Object -TypeName 'Microsoft.SqlServer.Management.SMO.Database' -ArgumentList $Server , $CopyDBName
    $CopyDB.Create() 
    $ObjTransfer   = New-Object -TypeName Microsoft.SqlServer.Management.SMO.Transfer -ArgumentList $SourceDB
    $ObjTransfer.CopyAllTables = $true
    $ObjTransfer.Options.WithDependencies = $true
    $ObjTransfer.Options.ContinueScriptingOnError = $true
    $ObjTransfer.DestinationDatabase = $CopyDBName
    $ObjTransfer.DestinationServer = $Server.Name
    $ObjTransfer.DestinationLoginSecure = $true
    $ObjTransfer.CopySchema = $true
    $ObjTransfer.ScriptTransfer()
    $ObjTransfer.TransferData()
}
function createRemoteContextServices {
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
                New-Service -Name "$remoteContextServiceName" -BinaryPathName "$remoteContextServiceBinaryPath" -DisplayName "$remoteContextServiceName" -StartupType Automatic -Description "$remoteContextServiceDescription" -Credential $global:NetworkServiceCredentials
                Start-Service -Name $remoteContextServiceName
            }
        }
    }
}
function cloneAppPoolsAndFiles {
    $ws="_WS"
    copy "IIS:\AppPools\TAP_$global:instanceName" "IIS:\AppPools\TAP_$global:newInstanceName"
    copy "IIS:\AppPools\TAP_$global:instanceName$ws" "iis:\apppools\TAP_$global:newInstanceName$ws"
    $physicalPathToNewInstance = $global:pathToInstance -replace "$global:instanceName", "$global:newInstanceName"
    copy $global:pathToInstance -Destination $physicalPathToNewInstance -Recurse -Force
    #Set-ItemProperty "IIS:\Sites\$global:siteName\$global:newInstanceName" applicationPool "TAP_$global:newInstanceName"
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
        echo $appPoolName
        Stop-WebAppPool $appPoolName
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
    (Get-Content -Path "$global:pathtoNewInstance\web.config") | ForEach-Object {$_ -replace ";DataBase=$global:SQLDBName;",";DataBase=$global:newInstanceName;"} | Set-Content -Path "$global:pathtoNewInstance\web.config"
    (Get-Content -Path "$global:pathtoNewInstance\Cache Server\VisualKPICache.xml") | ForEach-Object {$_ -replace ";DataBase=$global:SQLDBName;",";DataBase=$global:newInstanceName;"} | Set-Content -Path "$global:pathtoNewInstance\Cache Server\VisualKPICache.xml"
    (Get-Content -Path "$global:pathtoNewInstance\Alerter\VisualKPIAlerter.xml") | ForEach-Object {$_ -replace ";DataBase=$global:SQLDBName;",";DataBase=$global:newInstanceName;"} | Set-Content -Path "$global:pathtoNewInstance\Alerter\VisualKPIAlerter.xml"
    $remotecontextFolders = Get-ChildItem "$global:pathToNewInstance\Remote Context Services"
    foreach($folder in $RemoteContextFolders) {
        "$folder"
         $FolderChild = get-childItem "$global:pathToNewInstance\Remote Context Services\$folder" 
         foreach($child in $FolderChild) {
            $childName = Get-ItemProperty -path "$global:pathToNewInstance\Remote Context Services\$folder\$child"  | Select-Object Name
            $childName = $childName -replace "@{Name=", ""
            $childName = $childName -replace "}", ""
            if($childName.endswith("Service.xml") ) {
             (Get-Content -Path "$global:pathtoNewInstance\Remote Context Services\$folder\$child") | ForEach-Object {$_ -replace ";DataBase=$global:SQLDBName;",";DataBase=$global:newInstanceName;"} | Set-Content -Path "$global:pathtoNewInstance\Remote Context Services\$folder\$child"
            }
         }
    }
}

function editDatabaseInterfacesURL { 
    $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
    $SqlConnection.ConnectionString = "Server = $global:SQLServer; Database = 
    $global:newInstanceName; Integrated Security = True"
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
try {

getInputs
if(hasInputs -eq 1){return}
#cloneIISApp #clones the file system stuff
$global:pathToInstance = Get-WebFilePath "iis:\Sites\$global:siteName\$global:instanceName"
stopAppPools
cloneAppPoolsAndFiles

$global:pathToNewInstance = Get-WebFilePath "iis:\Sites\$global:siteName\$global:newInstanceName"
createRemoteContextServices
createBaseServices
replaceConnectionStrings
startAppPools
"We will now clone the database, please standby"
cloneDatabase
editDatabaseInterfacesURL

} catch {"This is embarrasing, an error has ocurred"}