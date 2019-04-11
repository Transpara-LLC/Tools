#TODO create services and Rcs
#clone app pools
#make web apps


#Enforce Running as an administrator
function Use-RunAs {    
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
#end of enforcing administrator


#imports
Import-Module webadministration
Import-Module SQLPS -DisableNameChecking
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")
[system.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic")
#endimports

"--------------------------------------Defining variables--------------------------------------"
#global hardcodable variables
$global:siteName="Default Web Site"
$global:sourceInstanceName="GoodInstance"
$global:targetFolder = ""
#end of global variables
















#----------------------dont edit variables below--------------



$global:pathToInstance=""#the path to the source instance
$global:sourceSQLDBName=""#will be obtained and set by looking on the file system webconfig
$global:SQLServer=""#will be obtained and set by looking on the file system webconfig
$global:remoteContextPath = ""#the path of the remote context services 
$global:remoteContextExists = ""# a variable to check if there are remote context services
$networkServicePassword = (new-object System.Security.SecureString)
$global:NetworkServiceCredentials = New-Object System.Management.Automation.PSCredential ("NT AUTHORITY\NETWORK SERVICE", $networkServicePassword)

"--------------------------------------Defining functions--------------------------------------"
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


#replaces the connection strings of the cloned files (for the instance, rcs, cache server and alerter)
function replaceConnectionStrings {
     appendTextToRestoreScript('(Get-Content -Path "$pathToTargetInstance\web.config") | ForEach-Object {$_ -replace ";DataBase=$global:sourceSQLDBName;",";DataBase=$global:targetSQLDBName;"} | Set-Content -Path "$pathToTargetInstance\web.config"')
     appendTextToRestoreScript('(Get-Content -Path "$pathToTargetInstance\Cache Server\VisualKPICache.xml") | ForEach-Object {$_ -replace ";DataBase=$global:sourceSQLDBName;",";DataBase=$global:targetSQLDBName;"} | Set-Content -Path "$pathToTargetInstance\Cache Server\VisualKPICache.xml"')
     appendTextToRestoreScript('(Get-Content -Path "$pathToTargetInstance\Alerter\VisualKPIAlerter.xml") | ForEach-Object {$_ -replace ";DataBase=$global:sourceSQLDBName;",";DataBase=$global:targetSQLDBName;"} | Set-Content -Path "$pathToTargetInstance\Alerter\VisualKPIAlerter.xml"')
    
    if(!$global:remoteContextExists){ #if the remote context folder doesnt exist then we stop here
        "No RCS"
        return
    }
    appendTextToRestoreScript('$remotecontextFolders = Get-ChildItem "$pathToTargetInstance\Remote Context Services"')
    appendTextToRestoreScript('foreach($folder in $RemoteContextFolders) {')
        appendTextToRestoreScript('"$folder"')
         appendTextToRestoreScript('$FolderChild = get-childItem "$pathToTargetInstance\Remote Context Services\$folder" ')
         appendTextToRestoreScript('foreach($child in $FolderChild) {')
            appendTextToRestoreScript('$childName = Get-ItemProperty -path "$pathToTargetInstance\Remote Context Services\$folder\$child"  | Select-Object Name')
            appendTextToRestoreScript('$childName = $childName -replace "@{Name=", ""')
            appendTextToRestoreScript('$childName = $childName -replace "}", ""')
            appendTextToRestoreScript('if($childName.endswith("Service.xml") ) {')
             appendTextToRestoreScript('(Get-Content -Path "$pathToTargetInstance\Remote Context Services\$folder\$child") | ForEach-Object {$_ -replace ";DataBase=$global:sourceSQLDBName;",";DataBase=$global:targetSQLDBName;"} | Set-Content -Path "$pathToTargetInstance\Remote Context Services\$folder\$child"')
            appendTextToRestoreScript('}')
         appendTextToRestoreScript('}')
    appendTextToRestoreScript('}')

}






function mapDatabaseRoleToLogin {
    #get variables
    appendTextToRestoreScript('$instanceName = "$global:targetSQLServer"')
    appendTextToRestoreScript('$loginName = "NT AUTHORITY\NETWORK SERVICE"')
    appendTextToRestoreScript('$dbUserName = "NT AUTHORITY\NETWORK SERVICE"')
    appendTextToRestoreScript('$databasename = $global:targetSQLDBName')
    appendTextToRestoreScript('$roleName = "db_owner"')
    appendTextToRestoreScript('$server = New-Object -TypeName Microsoft.SqlServer.Management.Smo.Server -ArgumentList $instanceName')
    appendTextToRestoreScript('$server.Logins.Refresh()')
    #$server.Logins | Format-Table -Property Parent, ID, Name, CreateDate,  LoginType
    #add a database mapping
    appendTextToRestoreScript('$database = $server.Databases[$databasename]')
    appendTextToRestoreScript('$login = $server.Logins[$loginName]')
    appendTextToRestoreScript('if ($database.Users[$dbUserName])')
    appendTextToRestoreScript('{')
    appendTextToRestoreScript('    $database.Users[$dbUserName].Drop()')
    appendTextToRestoreScript('}')
    appendTextToRestoreScript('$dbUser = New-Object `')
    appendTextToRestoreScript('-TypeName Microsoft.SqlServer.Management.Smo.User `')
    appendTextToRestoreScript('-ArgumentList $database, $loginName')
    appendTextToRestoreScript('$dbUser.Login = $loginName')
    appendTextToRestoreScript('$dbUser.Create()')

    #assign database role for a new user
    appendTextToRestoreScript('$dbrole = $database.Roles[$roleName]')
    appendTextToRestoreScript('$dbrole.AddMember($dbUserName)')
    appendTextToRestoreScript('$dbrole.Alter')
}

function appendTextToRestoreScript($text) {
    Add-Content "$global:targetFolder\restore.ps1" "$text"
}#adds text to the restore.ps1 file


#returns 1 if an input required to execute the program is missing. It is used to stop the execution
function hasInputs {
    if($global:siteName -ne "" -and $global:sourceInstanceName -ne "" -and $global:targetFolder -ne ""  ) {
         return 0
    }
    return 1    
}
function getAppPoolIdentity($appPoolName) {
   return (get-item IIS:\AppPools\$appPoolName).processModel.identityType
}
function getAppPoolValue($appPoolName) {
    if(getAppPoolIdentity($appPoolName) -eq "NetworkService") {
        return 2
    }
}

function packDB {
    $log="_log"
    Backup-SqlDatabase -ServerInstance "$global:SQLServer" -Database "$global:sourceSQLDBName" -Initialize
    [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.SMO') | out-null
    $s = New-Object ('Microsoft.SqlServer.Management.Smo.Server') "$global:SQLServer" 
    $backupDirectoryPath= $s.Settings.BackupDirectory
    $DBBackupPath = "$backupDirectoryPath\$sourceSQLDBName.bak"
    copy "$DBBackupPath" "$targetFolder"

    $at = '"@'
    $db_path = '$DBBackupPath'
    $sc_sql_name = '$global:sourceSQLDBName'
    $trg_sql_name = '$global:targetSQLDBName'
    $dt_fl_location = '$dataFileLocation'
    $lg_fl_location = '$logFileLocation' 
    $log='$log'
    appendTextToRestoreScript('$log="_log"')
    appendTextToRestoreScript('[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.SMO") | out-null')
    appendTextToRestoreScript('$s = New-Object ("Microsoft.SqlServer.Management.Smo.Server") "$global:targetSQLServer"') 
    appendTextToRestoreScript('$MDFDirectory = $s.Settings.DefaultFile')
    appendTextToRestoreScript('$LDFDirectory = $s.Settings.DefaultLog')

    appendTextToRestoreScript('$DBBackupPath = "$PSScriptRoot\$sourceSQLDBName.bak"')
    appendTextToRestoreScript('$dataFileLocation ="$MDFDirectory\$global:targetSQLDBName.mdf"')
    appendTextToRestoreScript('$logFileLocation = "$LDFDirectory\$global:targetSQLDBName$log.ldf"')
    appendTextToRestoreScript('$sql = @"')
    appendTextToRestoreScript('select user_name(),suser_sname()')
    appendTextToRestoreScript('GO')  
    appendTextToRestoreScript('USE [master]')
    appendTextToRestoreScript('RESTORE DATABASE [$global:targetSQLDBName]') 
    appendTextToRestoreScript("FROM DISK = N'$db_path'") 
    appendTextToRestoreScript('WITH FILE = 1,')  
    appendTextToRestoreScript("    MOVE N'$sc_sql_name' TO N'$dt_fl_location',")  
    appendTextToRestoreScript("    MOVE N'$sc_sql_name$log' TO N'$lg_fl_location',")  
    appendTextToRestoreScript('    NOUNLOAD, REPLACE, STATS = 5')
    appendTextToRestoreScript('GO')
    appendTextToRestoreScript("ALTER DATABASE [$trg_sql_name] MODIFY FILE ( NAME = '$sc_sql_name', NEWNAME = '$trg_sql_name');")
    appendTextToRestoreScript("ALTER DATABASE [$trg_sql_name] MODIFY FILE ( NAME = '$sc_sql_name$log', NEWNAME = '$trg_sql_name$log');")
    appendTextToRestoreScript("ALTER DATABASE [$trg_sql_name] SET MULTI_USER") 
    appendTextToRestoreScript($at)

    appendTextToRestoreScript('invoke-sqlcmd $sql -ServerInstance $global:targetSQLServer  ')
   
}

function packSystemFilesAndSite {
    copy $global:pathToInstance -Destination $global:targetFolder -Recurse -Force #copies the files
    appendTextToRestoreScript('$global:sourceInstanceName="'+$global:sourceInstanceName+'"')
    appendTextToRestoreScript('$pathToTargetSite=Get-WebFilePath "iis:\Sites\$global:targetSiteName\"')
    appendTextToRestoreScript('copy $PSScriptRoot\$global:sourceInstanceName -Destination $pathToTargetSite\$global:targetInstanceName -Recurse -Force')
    appendTextToRestoreScript('$pathToTargetInstance=Get-WebFilePath "iis:\Sites\$global:targetSiteName\$global:targetInstanceName"')
    $ws="_WS" 
    appendTextToRestoreScript('$ws="_WS"')
    appendTextToRestoreScript('$mainAppPoolName = "TAP_$global:targetInstanceName"')
    appendTextToRestoreScript('$wsAppPoolName =  "TAP_$global:targetInstanceName$ws"')
    $oldMainAppPoolName = "TAP_$global:sourceInstanceName"
    $oldWsAppPoolName = "TAP_$global:sourceInstanceName$ws"
    $mainAppPoolIdentityValue = getAppPoolValue($oldMainAppPoolName)
    $wsAppPoolIdentityValue = getAppPoolValue($oldWsAppPoolName)
    appendTextToRestoreScript('New-WebAppPool $mainAppPoolName')
    appendTextToRestoreScript('New-WebAppPool $wsAppPoolName')
    appendTextToRestoreScript('Set-ItemProperty "IIS:\AppPools\$mainAppPoolName" -name processModel.identityType -value '+$mainAppPoolIdentityValue)
    appendTextToRestoreScript('Set-ItemProperty "IIS:\AppPools\$wsAppPoolName" -name processModel.identityType -value '+$wsAppPoolIdentityValue)

    appendTextToRestoreScript('ConvertTo-WebApplication "IIS:\Sites\$global:targetSiteName\$global:targetInstanceName" -ApplicationPool "TAP_$global:targetInstanceName"')
    appendTextToRestoreScript('ConvertTo-WebApplication "IIS:\Sites\$global:targetSiteName\$global:targetInstanceName\Interfaces" -ApplicationPool "TAP_$global:targetInstanceName"')
    appendTextToRestoreScript('ConvertTo-WebApplication "IIS:\Sites\$global:targetSiteName\$global:targetInstanceName\Webservice" -ApplicationPool "TAP_$global:targetInstanceName$ws"')
    
    $interfaces = Get-ChildItem "iis:\Sites\$global:siteName\$global:sourceInstanceName\Interfaces"
    foreach($interface in $interfaces) {
        $interfaceName = Get-ItemProperty -path "$interface"  | Select-Object Name
        $interfaceName = $interfaceName -replace "@{Name=", ""
        $interfaceName = $interfaceName -replace "}", ""
        $appPoolName = Get-ItemProperty -path "iis:\Sites\$siteName\$global:sourceInstanceName\Interfaces\$interfaceName" | Select-Object applicationPool 
        $appPoolName = $AppPoolName -replace "@{applicationPool=", ""
        $appPoolName = $AppPoolName -replace "}", ""
        appendTextToRestoreScript('$interfaceName= "'+$interfaceName+'"')
        appendTextToRestoreScript('$appPoolName= "'+$appPoolName+'"')
        appendTextToRestoreScript('$newAppPoolName = $appPoolName -replace "$global:sourceInstanceName","$global:targetInstanceName"')
        $newAppPoolIdentityValue = getAppPoolValue($appPoolName)
        appendTextToRestoreScript('New-WebAppPool $newAppPoolName')
        appendTextToRestoreScript('Set-ItemProperty "IIS:\AppPools\$newAppPoolName" -name processModel.identityType -value '+$newAppPoolIdentityValue)
        appendTextToRestoreScript('ConvertTo-WebApplication "IIS:\Sites\$global:targetSiteName\$global:targetInstanceName\Interfaces\$interfaceName" -ApplicationPool $newAppPoolName') 

    }
}


function getInputs { 
    if($global:siteName -eq "") {
        $global:siteName = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the name of the site where the VKPI instance is", "Site Name Request")
    }
    if($global:sourceInstanceName -eq "") {   
         $global:sourceInstanceName =  [Microsoft.VisualBasic.Interaction]::InputBox("Enter the old instance name", "Instance Name Request")
    }
    if($global:targetFolder -eq "") {   
         [System.Windows.Forms.MessageBox]::Show('Select the folder for the package', 'Information', 'Ok', 'Information')
         $global:targetFolder =  Get-Folder
    }
}#gets all the user inputs if they are not already hardcoded, if the new SQL database name is not entered, it will be the same as the new instance name
Function Get-Folder($initialDirectory) { 
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")|Out-Null

    $foldername = New-Object System.Windows.Forms.FolderBrowserDialog
    $foldername.Description = "Select a folder"
    $foldername.rootfolder = "MyComputer"

    if($foldername.ShowDialog() -eq "OK")
    {
        $folder += $foldername.SelectedPath
    }
    return $folder
}#Prompts a GUI for selecting a folder and returns the path

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
      appendTextToRestoreScript('$sourceSQLDBName="'+$global:sourceSQLDBName+'"')
    }
  }
}
function editDatabaseInterfacesURL { 
    appendTextToRestoreScript('$SQLTable = "dbo.tableInterfaces" #the table to edit the interfaces url')
    appendTextToRestoreScript('$SqlConnection = New-Object System.Data.SqlClient.SqlConnection')
    appendTextToRestoreScript('$SqlConnection.ConnectionString = "Server = $global:targetSQLServer; Database =') 
    appendTextToRestoreScript('$global:targetSQLDBName; Integrated Security = True"')
    appendTextToRestoreScript('$SqlConnection.Open()')
    appendTextToRestoreScript('$update = @"')
     appendTextToRestoreScript('UPDATE $SQLTable')
     appendTextToRestoreScript('SET') 
     $src = '$global:sourceInstanceName'
     $trg = '$global:targetInstanceName'
     $line = "[URL] = REPLACE([URL], '$src/interfaces/', '$trg/interfaces/')"
     appendTextToRestoreScript($line)
appendTextToRestoreScript('"@')
    appendTextToRestoreScript('$dbwrite = $SqlConnection.CreateCommand()')
    appendTextToRestoreScript('$dbwrite.CommandText = $update')
    appendTextToRestoreScript('$dbwrite.ExecuteNonQuery()') 
    appendTextToRestoreScript('$Sqlconnection.Close()')
    appendTextToRestoreScript('"Edited Interface SQL Table"')
}

function createBaseServices {
    appendTextToRestoreScript('$newCacheServiceName="Visual KPI Cache Server ($global:targetInstanceName)"')
    appendTextToRestoreScript('$newAlertServiceName="Visual KPI Alert Server ($global:targetInstanceName)"')
    $oldCacheName  ="Visual KPI Cache Server ($global:sourceInstanceName)"
    $oldAlertName = "Visual KPI Alert Server ($global:sourceInstanceName)"
    $cacheStartType = getServiceStartupType($oldCacheName)
    $alertStartType = getServiceStartupType($oldAlertName)
    appendTextToRestoreScript('New-Service -Name "$newCacheServiceName" -BinaryPathName "$global:pathToTargetInstance\Cache Server\VisualKPICache.exe" -DisplayName "$newCacheServiceName" -StartupType '+$cacheStartType+' -Description "Enables caching of KPI data in Visual KPI Server" -Credential $global:NetworkServiceCredentials')
    appendTextToRestoreScript('New-Service -Name "$newAlertServiceName" -BinaryPathName "$global:pathToTargetInstance\Alerter\VisualKPIAlerter.exe" -DisplayName "$newAlertServiceName" -StartupType '+$alertStartType+' -Description "Enables sending of Alerts based on KPI state changes in Visual KPI Server" -Credential $global:NetworkServiceCredentials')
    $cacheStatus=getServiceStatus($oldCacheName)
    $alertStatus=getServiceStatus($oldAlertName)
    if($cacheStatus -eq "Running") {
         appendTextToRestoreScript('Start-Service -Name $newCacheServiceName')
    }
    if($alertStatus -eq "Running") {
         appendTextToRestoreScript('Start-Service -Name $newAlertServiceName')
    } 
}

function createRemoteContextServices {
    if(!$global:remoteContextExists){
        "No RCS"
        return
    }
    appendTextToRestoreScript('$RemoteContextFolders = Get-ChildItem "$global:pathToTargetInstance\Remote Context Services"')
    appendTextToRestoreScript('foreach($folder in $RemoteContextFolders) {')
        appendTextToRestoreScript('"$folder"')
        appendTextToRestoreScript('$FolderChild = get-childItem "$global:pathToTargetInstance\Remote Context Services\$folder"') 
        appendTextToRestoreScript('foreach($child in $FolderChild) {')
            appendTextToRestoreScript('$childName = Get-ItemProperty -path "$global:pathToTargetInstance\Remote Context Services\$folder\$child"  | Select-Object Name')
            appendTextToRestoreScript('$childName = $childName -replace "@{Name=", ""')
            appendTextToRestoreScript('$childName = $childName -replace "}", ""')
            appendTextToRestoreScript('if($childName.endswith("Service.exe") ) {')
                appendTextToRestoreScript('$folderName = Get-ItemProperty -path "$global:pathToTargetInstance\Remote Context Services\$folder"  | Select-Object Name')
                appendTextToRestoreScript('$folderName = $folderName -replace "@{Name=", ""')
                appendTextToRestoreScript('$folderName = $folderName -replace "}", ""')
                appendTextToRestoreScript('$serviceExecutableName = $childName')
                appendTextToRestoreScript('$remoteContextServiceName = ""')
                appendTextToRestoreScript('$remoteContextServiceBinaryPath= ""')
                appendTextToRestoreScript('$remoteContextServiceDescription =""')
                appendTextToRestoreScript('if($serviceExecutableName.contains("SQL")) {')
                    appendTextToRestoreScript('$remoteContextServiceName = "Visual KPI SQL RC Server ($global:targetInstanceName : $folderName)"')
                    appendTextToRestoreScript('$remoteContextServiceDescription ="Enables automatic creation of Visual KPI Components from a source SQL database"')
                appendTextToRestoreScript('} elseif($serviceExecutableName.contains("NIWNAS")) {')
                    appendTextToRestoreScript('$remoteContextServiceName = "Visual KPI NIWNAS RC Server ($global:targetInstanceName : $folderName)"')
                    appendTextToRestoreScript('$remoteContextServiceDescription ="Enables automatic creation of Visual KPI Components from a NIWNAS server"')
                appendTextToRestoreScript('} elseif($serviceExecutableName.contains("Monday")) {')
                    appendTextToRestoreScript('$remoteContextServiceName = "Visual KPI Monday RC Server ($global:targetInstanceName : $folderName)"')
                    appendTextToRestoreScript('$remoteContextServiceDescription ="Enables automatic creation of Visual KPI Components from a Monday.com site"')
                appendTextToRestoreScript('} elseif($serviceExecutableName.contains("inmation")) {')
                    appendTextToRestoreScript('$remoteContextServiceName = "Visual KPI inmation RC Server ($global:targetInstanceName : $folderName)"')
                    appendTextToRestoreScript('$remoteContextServiceDescription ="Enables automatic creation of Visual KPI Components from inmation Models"')
                appendTextToRestoreScript('} elseif($serviceExecutableName.contains("AFIntegration")) {')
                    appendTextToRestoreScript('$remoteContextServiceName = "Visual KPI AF RC Server ($global:targetInstanceName : $folderName)"')
                    appendTextToRestoreScript('$remoteContextServiceDescription ="Enables automatic creation of Visual KPI Components from AF Elements"')
                appendTextToRestoreScript('}')
                appendTextToRestoreScript('$remoteContextServiceBinaryPath= "$global:pathToTargetInstance\Remote Context Services\$folder\$child"')
                #$oldServiceName = $remoteContextServiceName -replace "$global:targetInstanceName", "$global:sourceInstanceName"') 
                #$startType = getServiceStartupType($oldServiceName)
                #$oldServiceStatus = getServiceStatus($oldServiceName)
                appendTextToRestoreScript('New-Service -Name "$remoteContextServiceName" -BinaryPathName "$remoteContextServiceBinaryPath" -DisplayName "$remoteContextServiceName" -StartupType "Automatic" -Description "$remoteContextServiceDescription" -Credential $global:NetworkServiceCredentials')
               # if($oldServiceStatus -eq "Running") {
                     appendTextToRestoreScript('Start-Service -Name $remoteContextServiceName')
    
                #}
              
            appendTextToRestoreScript('}')
        appendTextToRestoreScript('}')
    appendTextToRestoreScript('}')
}




getInputs
if(hasInputs -eq 1){return}

$global:pathToInstance = Get-WebFilePath "iis:\Sites\$global:siteName\$global:sourceInstanceName"

"--------------------------------------Getting server and database--------------------------------------"
appendTextToRestoreScript("Import-Module webadministration")
appendTextToRestoreScript("Import-Module SQLPS -DisableNameChecking")
appendTextToRestoreScript('[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")')
appendTextToRestoreScript('[system.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic")')
appendTextToRestoreScript('$networkServicePassword = (new-object System.Security.SecureString)')
appendTextToRestoreScript('$global:NetworkServiceCredentials = New-Object System.Management.Automation.PSCredential ("NT AUTHORITY\NETWORK SERVICE", $networkServicePassword)')

appendTextToRestoreScript('$global:targetSiteName = ""')
appendTextToRestoreScript('$global:targetInstanceName =""')
appendTextToRestoreScript('$global:targetSQLDBName = ""')
appendTextToRestoreScript('$global:targetSQLServer = ""')

appendTextToRestoreScript('$global:targetSiteName = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the name of the site where the VKPI instance will be", "Site Name Request")')
appendTextToRestoreScript('$global:targetInstanceName = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the name of the new instance", "Instance Name Request")')
appendTextToRestoreScript('$global:targetSQLDBName = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the name of the new SQL DB ", "SQL DB Name Request")')
appendTextToRestoreScript('$global:targetSQLServer = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the name of the new SQL SERVER", "SQL Server Name Request")')


appendTextToRestoreScript('if($global:targetSQLDBName -eq "") {')
appendTextToRestoreScript('$global:targetSQLDBName = $global:targetInstanceName}')
appendTextToRestoreScript('if($global:targetSiteName -eq "" -or $global:targetInstanceName -eq "" -or $global:targetSQLServer -eq ""){return}')



getServerAndDB

"--------------------------------------Stoping app pools--------------------------------------"
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
        if($appPoolName.contains("$global:sourceInstanceName")) {
            if((Get-WebAppPoolState -Name $appPoolName).Value -eq "Running") {
                $cango = ""
            }
        }
    }
}
until ($cango -eq "go")
"--------------------------------------App Pools stopped--------------------------------------"
$global:remoteContextPath = "$global:pathToInstance\Remote Context Services" 
$global:remoteContextExists = Test-Path -LiteralPath $remoteContextPath

packSystemFilesAndSite
packDB
createBaseServices
createRemoteContextServices
editDatabaseInterfacesURL
mapDatabaseRoleToLogin
replaceConnectionStrings
startAppPools