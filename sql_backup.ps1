########################################################################
# Program Name: mssql_backup_automation
# Description: This utility remotely backs up user databases on target database servers and uploads the backups to a destination FTP server
# Version: 1.0
# Completion Date: 3/27/2015
########################################################################

# Switch to enable silent mode
param([string]$SilentPassed="")
$global:Silent = $false
if ($SilentPassed -eq "Silent") {
    $global:Silent = $true
}

# Declare global variables
#- Script path when executed from a PowerShell console
$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition

#- Global constant variables
$PROGRAM_NAME = "mssql_backup_automation"
$global:Key = (1,4,2,3,56,34,254,222,1,1,2,23,42,54,33,233,1,34,2,7,6,5,35,43) # this key should match the key used in encryption
$global:LOGPATH = $scriptpath + "\Log\" + ("mssql_backup_log" + (date -Format "yyyyMMddhhmmss") + ".log")
$global:EMAILLOGPATH = $scriptpath + ("\mssql_backup_log" + (date -Format "yyyyMMddhhmmss") + ".log")

#- Default Parameters
#- These are backup values in case config.ini is not properly filled
#- This function can be ignored as long as config.ini is filled in properly
function Load-Defaults {
    $global:serverInstance="localhost"
    $global:timeout = 600
    $global:databases = @()
    $global:dbExclusions = @()
    $global:filteredDBs = @()
    $global:backupTempPath = "E:\Temp\DBBackup"
    $global:backupTempPathLocal = "$scriptPath\TEMP"
    $global:backupTempZIPPathLocal = "$scriptPath\TEMP_ZIP"
    $global:serverInstance = ""
    $global:errorEncountered = $false
    $global:ftpserver="<ftp_server_hostname>"
    $global:username=""
    $global:password=""
    $global:ftppath="<ftp_server_virtual_path>"
    $global:emailto="<you_email_address>"
    $global:emailcc="<you_coworker_email>"
    $global:emailfrom="MSQL Backup Automation <noreply@sqlbackupautomation.com>"
    $global:smtpserver="<ftp_server>"
    $global:smtpport=25
    $global:emailsubject="Backup of MSSQL databases have been completed"
    $global:emailerrorsubject="[ALERT] Backup of MSSQL databases failed"
    $global:emailbody=""
    $global:loghistory=20
    $global:dbList = ""
    $global:7zPath = "$scriptPath\Res\7z\7z.exe"
}

#- SQL commands
$COMMAND_LISTDBS = "SELECT * FROM SYS.DATABASES"
$COMMAND_LISTDBS += " WHERE DATABASE_ID > 4"

$COMMAND_BACKUP = "BACKUP DATABASE <database>"
$COMMAND_BACKUP += " TO DISK = N`'<backup_path>`'"
$COMMAND_BACKUP += " WITH COPY_ONLY, COMPRESSION, STATS=10"

$COMMAND_VERIFY = "RESTORE VERIFYONLY"
$COMMAND_VERIFY += " FROM DISK = `'<backup_path>`'"
$COMMAND_VERIFY += " WITH STATS=10"
##

# Main Program
function Main-Program 
{
    Output-Message -NoLog "$PROGRAM_NAME has been initiated"
    try {
        # Load SQLPS PowerShell Module
        Output-Message -NoLog "Loading SQLPS PowerShell Module"
        Import-Module -Name "$scriptPath\Res\sqlps" -DisableNameChecking > $null
        $VerbosePreference = "Continue"
        Output-Message -NoLog "SQLPS successfully loaded"

        # Get Database Servers information from configuration file(s)
        $configFiles = Get-ChildItem -Path "$scriptPath"  | Where-Object {$_.Name -like "*config*" -and $_.Name -like "*.ini*"} | Select-Object -ExpandProperty Fullname
        
        # Process each configuration file
        foreach ($configFile in $configFiles) {
            # Get Parameters
            $programStartTime = date
            Output-Message -NoLog "Loading configuration parameters"
            Load-Defaults
            $loadResult = Load-Settings -configFile $configFile
            Output-Message "$PROGRAM_NAME has been initiated on $global:serverInstance"

            # Get DB List
            if ($loadResult -eq $true) {
                Output-Message "Successfully loaded configuration data"
                Output-Message "Listing all USER databases in $global:serverInstance"
                # Setup query parameters
                $params = @{'serverInstance'=$global:serverInstance;
                            'Query'=$COMMAND_LISTDBS;
                            'QueryTimeout'=$global:timeout;
                           }
                $getListResult = LIST_DB @params
            }
            else {
                Output-Message -Red "Configuration load error was encountered"
                Output-Message -Red "[ERROR]: Unable to read from configuration file $configFile"
                $global:errorEncountered = $true
            }

            # Backup all USER databases (except exclusions) in database server
            if ($loadResult -eq $true -and $getListResult -eq $true) {
                Output-Message -Green "Successfully listed all databases from $global:serverInstance"
                Output-Message "Backing up each database in $global:serverInstance"
                # recurse through each Database in Database List
                $successbackup = $true
                foreach ($database in $global:databases) {
                    # Backup database
                    Output-Message "Backing up database $database in $global:serverInstance"
                    if ($successbackup -eq $true) {
                        # Setup query parameters
                        $params = @{'serverInstance'=$global:serverInstance;
                                    'Query'=$COMMAND_BACKUP;
                                    'QueryTimeout'=$global:timeout;
                                    'Database'=$database;
                                    'Backup_Path'=$global:backupTempPath;
                                   }
                        $backupDBResult = BACKUP_DB @params
                        if ($backupDBResult -eq $true) {
                            Output-Message "Database $database has been successfully backed up"
                            $successbackup = $true 
                        }
                        else {
                            Output-Message -Red "Unable to backup $database in $global:serverInstance"
                            Output-Message -Red "[ERROR]: $backupDBResult"
                            $successbackup = $false
                            $global:errorEncountered = $true
                        }
                    }
                }
            }
            else {
                Output-Message -Red "Unable to obtain list of databases in $global:serverInstance"
                Output-Message -Red "[ERROR]: $getListResult"
                $global:errorEncountered = $true
            }

            # Verify all database backups
            if ($loadResult -eq $true -and $getListResult -eq $true -and $backupDBResult -eq $true) {
                Output-Message -Green "Successfully backed up all USER databases from $global:serverInstance [Except exclusions]"
                Output-Message "Verifying backup files"
                # recurse through each Database in Database List
                $verifySuccess = $true
                foreach ($database in $global:databases) {
                    # Verify database backup
                    Output-Message "Verifying database backup for $database"
                    if ($verifySuccess -eq $true) {
                        # Setup query parameters
                        $params = @{'serverInstance'=$global:serverInstance;
                                    'Query'=$COMMAND_VERIFY;
                                    'QueryTimeout'=$global:timeout;
                                    'Database'=$database;
                                    'Backup_Path'=$global:backupTempPath;
                                   }
                        $verifybackupResult = VERIFY_BACKUP @params
                        if ($verifybackupResult -eq $true) {
                            Output-Message "Database backup $database.bak has been verified to be valid"
                            $verifySuccess = $true 
                        }
                        else {
                            Output-Message -Red "Unable to verify $database.bak"
                            Output-Message -Red "[ERROR]: $verifybackupResult"
                            $verifySuccess = $false
                            $global:errorEncountered = $true
                        }
                    }
                }
            }
            else {
                Output-Message -Red "Unable to back up all listed databases in $global:serverInstance"
                Output-Message -Red "[ERROR]: $backupDBResult"
                $global:errorEncountered = $true
            }

            # Get all Database Backups
            if ($loadResult -eq $true -and $getListResult -eq $true -and $backupDBResult -eq $true -and $verifybackupResult -eq $true) {
                Output-Message -Green "Successfully verified all USER database backups"
                Output-Message "Obtaining all database backups from $global:serverInstance to local temporary path $global:backupTempPathLocal"
                # recurse through each Database in Database List
                $getSuccess = $true
                foreach ($database in $global:databases) {
                    # Get database backup
                    Output-Message "Obtaining database backup $database.bak"
                    if ($getSuccess -eq $true) {
                        # Setup query parameters
                        $params = @{'serverInstance'=$global:serverInstance;
                                    'Database'=$database;
                                    'Backup_Path'=$global:backupTempPath;
                                    'Backup_Path_Local'=$global:backupTempPathLocal;
                                   }
                        $getBackupResult = Get-BackupFile @params
                        if ($getBackupResult -eq $true) {
                            Output-Message "$database.bak has been successfully transferred to temporary path"
                            $getSuccess = $true 
                        }
                        else {
                            Output-Message -Red "Unable to obtain $database.bak"
                            Output-Message -Red "[ERROR]: $getBackupResult"
                            $getSuccess = $false
                            $global:errorEncountered = $true
                        }
                    }
                }
            }
            else {
                Output-Message -Red "Unable to verify all database backups from $global:serverInstance"
                Output-Message -Red "[ERROR]: $verifybackupResult"
                $global:errorEncountered = $true
            }

            # ZIP all Database Backups
            if ($loadResult -eq $true -and $getListResult -eq $true -and $backupDBResult -eq $true -and $verifybackupResult -eq $true -and $getBackupResult -eq $true) {
                Output-Message -Green "Successfully obtained all USER database backups"
                Output-Message "Compressing all backups to ZIP archives"
                if ((Test-Path $global:backupTempZIPPathLocal) -eq $false) {
                    Output-Message "Creating temporary archive directory $global:backupTempZIPPathLocal"
                    New-Item -Path $global:backupTempZIPPathLocal -ItemType Directory -Force -ErrorAction SilentlyContinue > $null
                }
                # recurse through each Database in Database List
                $zipSuccess = $true
                foreach ($database in $global:databases) {
                    # ZIP database backup
                    Output-Message "Zipping database backup $database.bak"
                    if ($zipSuccess -eq $true) {
                        # Setup query parameters
                        $params = @{'InputFile'=($global:backupTempPathLocal + "\$database.bak");
                                    'OutputFile'=($global:backupTempZIPPathLocal + "\$database.zip");
                                   }
                        $getZIPResult = ZIP-BackupFile @params
                        if ($getZIPResult -eq $true) {
                            Output-Message "$database.zip has been successfully created"
                            $zipSuccess = $true
                        }
                        else {
                            Output-Message -Red "Unable to create $database.zip"
                            Output-Message -Red "[ERROR]: $getZIPResult"
                            $zipSuccess = $false
                            $global:errorEncountered = $true
                        }
                    }
                }
                if ($zipSuccess -eq $true) {
                    # Drop Trigger File
                    Output-Message "Appending log file to triger file"
                    New-Item -ItemType File -Path "$global:backupTempZIPPathLocal\Trigger.txt" -Force -ErrorAction SilentlyContinue > $null
                    Copy-Item -Path $global:LOGPATH -Destination "$global:backupTempZIPPathLocal\Trigger.txt" -Force -ErrorAction SilentlyContinue > $null
                }
            }
            else {
                Output-Message -Red "Unable to obtain all database backups from $global:serverInstance"
                Output-Message -Red "[ERROR]: $getBackupResult"
                $global:errorEncountered = $true
            }

            # Transfer all obtained backups to FTP server
            if ($loadResult -eq $true -and $getListResult -eq $true -and $backupDBResult -eq $true -and $verifybackupResult -eq $true -and $getZIPResult -eq $true) {
                Output-Message -Green "Successfully compressed all USER database backups to ZIP archives"
                Output-Message "Transferring database backups to $global:ftpServer ..."
                Output-Message -Yellow "[DISCLAIMER]: Backup files may take some time to upload; the console will not show the upload progress"
                $ftpResult = FTP-BackupFiles # Parameters already set as global parameters
                if ($ftpResult -eq $true) {
                    Output-Message -Green "Successfully transferred all database backups to $global:ftpServer"
                }
                else {
                    Output-Message -Red "Unable to transfer all database backups to $global:ftpServer"
                    Output-Message -Red "[ERROR]: $ftpResult"
                    $global:errorEncountered = $true
                }
            }

            $programStopTime = date
            #Get Execution Duration
            $rawDuration = $programStopTime - $programStartTime
            if ($rawDuration.hours -ne 0)
            {
                $global:duration = $rawDuration.hours.tostring() + " Hours "
            }
            if ($rawDuration.minutes -ne 0)
            {
                $global:duration += $rawDuration.minutes.tostring() + " Minutes " 
            }
            if ($rawDuration.seconds -ne 0)
            {
                $global:duration += $rawDuration.seconds.tostring() + " Seconds"
            }
            ##

            # Email Result
            Output-Message "Sending email report to specified recipients"
            if ($ftpResult -eq $true -and $global:errorEncountered -eq $false) {
                $emailResult = Email-Result
            }
            else {
                $emailResult = Email-Result -Error
            }
            if ($emailResult -eq $true) {
                Output-Message -Green "Successfully sent email to specified recipients"
            }
            else {
                Output-Message -Red "Unable to send email to specified recipients"
                Output-Message -Red "[ERROR]: $emailResult"
                $global:errorEncountered = $true
            }

            # Cleanup
            Output-Message "Performing cleanup of temporary files"
            $params = @{'serverInstance'=$global:serverInstance;
                        'Backup_Path'=$global:backupTempPath;
                        'Backup_Path_Local'=$global:backupTempPathLocal;
                       }
            $cleanResult = Clean-TempFiles @params
            if ($cleanResult -eq $true) {
                $cleanLogs = Clean-OldLogs
                if ($cleanLogs -eq $true) {
                    Output-Message "Cleanup complete"
                }
                else {
                    Output-Message "[WARNING]: Unable to remove some log files"
                    Output-Message -Silent "[Return Message]: $cleanLogs"
                }
            }
            else {
                Output-Message "[WARNING]: Unable to remove one or more temporary files"
                Output-Message -Silent "[Return Message]: $cleanResult"
            }
        }
        # Prompt success or failure notification
        if ($Silent -eq $false) {
            [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") > $null
            if ($global:errorEncountered -eq $false) {
                [System.Windows.Forms.MessageBox]::Show("SQL Server Backup Automation completed successfully", "SUCCESS", 0, "Information") > $null
            }
            else {
                [System.Windows.Forms.MessageBox]::Show("One or more errors have been encountered`nRefer to the log file for details", "ERROR", 0, "Error") > $null
                & $global:LOGPATH
            }
        }
    }
    catch [Exception] {
        Output-Message -Red "Runtime error encountered during program execution"
        Output-Message -Red "[ERROR]: $_"
    }
}

function Load-Settings
{
    [CmdletBinding()]
    param(
    [parameter()]
    [string]$configFile)

    if ((Test-Path $configFile) -eq $true)
    {
        $configData = Get-Content -Path $configFile
        foreach ($line in $configData)
        {
            if ($line -ne "" -and $line -notlike "#*") {
                if ($line -like "*ServerInstance*") {
                    $global:serverInstance = $line.Split('=')[1]
                    $global:LOGPATH = $global:LOGPATH -replace "mssql_backup_log_",("SQL_Server_Backup_LOG_$global:serverInstance" + "_")
                    $global:EMAILLOGPATH = $global:EMAILLOGPATH -replace "mssql_backup_log_",("SQL_Server_Backup_LOG_$global:serverInstance" + "_")
                }
                if ($line -like "*Backup_Path*") {
                    $global:backupTempPath = $line.Split('=')[1]
                }
                if ($line -like "*Backup_Temp_Path*") {
                    if ($line -like "*E:*" -or $line -like "*C:*") {
                        $global:backupTempPathLocal = $line.Split('=')[1]
                    }
                    else {
                        $global:backupTempPathLocal = $scriptpath + '\' + $line.Split('=')[1]
                    }
                }
                if ($line -like "*Query_Timeout*") {
                    $global:timeout = $line.Split('=')[1]
                }
                if ($line -like "*DB_Exclude*") {
                    if ($line.Split('=')[1] -ne "" -and $line.Split('=')[1] -ne $null) {
                        $global:dbExclusions += $line.Split('=')[1]
                    }
                }
                if ($line -like "*Backup_Server*") {
                    $global:ftpserver = $line.Split('=')[1]
                }
                if ($line -like "*Backup_Virtual_Path*") {
                    $global:ftppath = $line.Split('=')[1]
                }
                if ($line -like "*FTP_Credential_File*") {
                    if ($line -like "*E:*" -or $line -like "*C:*") {
                        $ftpFilePath = $line.Split('=')[1]
                    }
                    else {
                        $ftpFilePath = $scriptpath + '\' + $line.Split('=')[1]
                    }
                    if (Test-Path $ftpFilePath)
                    {
                        $global:username = Get-DecryptedFTPUser -DecryptKey $global:Key -Path $ftpFilePath
                        $global:password = Get-DecryptedFTPPassword -DecryptKey $global:Key -Path $ftpFilePath
                    }
                    else {
                        return "FTP credential file not found"
                    }
                }
                if ($line -like "*Email_To*") {
                    $global:emailto = $line.Split('=')[1]
                }
                if ($line -like "*Email_CC*") {
                    $global:emailcc = $line.Split('=')[1]
                }
                if ($line -like "*Email_From*") {
                    $global:emailfrom = $line.Split('=')[1]
                }
                if ($line -like "*Email_SMTP_Server*") {
                    $global:smtpserver = $line.Split('=')[1]
                }
                if ($line -like "*Email_SMTP_Port*") {
                    $global:smtpport = $line.Split('=')[1]
                }
                if ($line -like "*Email_Subject*") {
                    $global:emailsubject = $line.Split('=')[1]
                }
                if ($line -like "*Email_Error_Subject*") {
                    $global:emailerrorsubject = $line.Split('=')[1]
                }
                if ($line -like "*Email_Body*") {
                    $global:emailbody += $line.Split('=')[1]
                    $global:emailbody += "`n"
                }
                if ($line -like "*Log_History*") {
                    $global:loghistory = $line.Split('=')[1]
                }
            }
        }
    }
    else {
        return "Configuration file not found"
    }
    return $true
}

function LIST_DB
{
    [CmdletBinding()]
    param(
    [parameter()]
    [string]$serverInstance,

    [parameter()]
    [string]$Query = $COMMAND_LISTDBS,

    [parameter()]
    [string]$QueryTimeout = $global:timeout
    )
    BEGIN {}
    PROCESS
    {
        try {
            Output-Message "Invoking remote SQL command on $serverInstance to get a list of all USER databases"
            $invokeResult = Invoke-Sqlcmd -Query $Query -ServerInstance $serverInstance -QueryTimeout $QueryTimeout -OutputSqlErrors $true | Select-Object -ExpandProperty Name
            Output-Message "Remote SQL process completed"
            if ($? -eq $true) {
                $dbList = @()
                if ($global:dbExclusions -eq $null) {
                    $global:databases = $invokeResult
                }
                else {
                    foreach ($result in $invokeResult) {
                        Output-Message "Processing database $result for DB Exclusion rule"
                        if ($global:dbExclusions -notcontains $result) {
                            Output-Message "Database $result marked for backup"
                            $dbList += $result
                        }
                        else {
                            # Notify Exclusion
                            Output-Message "$result filtered out due to DB Exclusion rule"
                            $global:filteredDBs += $result
                        }
                    }
                    $global:databases = $dbList
                }
                return $true
            }
            else {
                return $Error[0].Exception
            }
        }
        catch [Exception] {
            return $_
        }
    }
    END {}
}

function BACKUP_DB
{
    [CmdletBinding()]
    param(
    [parameter()]
    [string]$serverInstance,

    [parameter()]
    [string]$Database,

    [parameter()]
    [string]$Backup_Path = $global:backupTempPath,

    [parameter()]
    [string]$Query = $COMMAND_BACKUP,

    [parameter()]
    [string]$QueryTimeout = $global:timeout
    )
    BEGIN {
        $Query = $Query -replace '<database>',$Database
        $Query = $Query -replace '<backup_path>',"$Backup_Path\$Database.bak"
    }
    PROCESS
    {
        try {
            # Create Temporary backup Path
            $Backup_Path_UNC = PPath-ToUNC -Server $serverInstance -PPath $Backup_Path
            if ((Test-Path -Path $Backup_Path_UNC) -eq $false) {
                New-Item -Path $Backup_Path_UNC -ItemType Directory -Force > $null
            }
            # Create backup
            Output-Message "Invoking remote SQL command on $serverInstance to backup $Database"
            Output-Message -Yellow -NoLog "[DISCLAIMER]: Some large databases may take some time to back up. Please be patient..."
            $invokeResult = Invoke-Sqlcmd -Query $Query -ServerInstance $serverInstance -QueryTimeout $QueryTimeout -OutputSqlErrors $true -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name
            Output-Message "Remote SQL process completed"
            if ($invokeResult -eq $null -and $? -eq $true) {
                return $true
            }
            else {
                return $Error[0].Exception
            }
        }
        catch [Exception] {
            return $_
        }
    }
    END {}
}

function VERIFY_BACKUP
{
    [CmdletBinding()]
    param(
    [parameter()]
    [string]$serverInstance,

    [parameter()]
    [string]$Database,

    [parameter()]
    [string]$Backup_Path = $global:backupTempPath,

    [parameter()]
    [string]$Query = $COMMAND_VERIFY,

    [parameter()]
    [string]$QueryTimeout = $global:timeout
    )
    BEGIN {
        $Query = $Query -replace '<database>',$Database
        $Query = $Query -replace '<backup_path>',"$Backup_Path\$Database.bak"
    }
    PROCESS
    {
        try {
            Output-Message "Invoking remote SQL command on $serverInstance to verify backup of $Database"
            $invokeResult = Invoke-Sqlcmd -Query $Query -ServerInstance $serverInstance -QueryTimeout $QueryTimeout -OutputSqlErrors $true -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name
            Output-Message "Remote SQL process completed"
            if ($invokeResult -eq $null -and $? -eq $true) {
                return $true
            }
            else {
                return $Error[0].Exception
            }
        }
        catch [Exception] {
            return $_
        }
    }
    END {}
}

function Get-BackupFile
{
    [CmdletBinding()]
    param(
    [parameter()]
    [string]$serverInstance,

    [parameter()]
    [string]$Database,

    [parameter()]
    [string]$Backup_Path = $global:backupTempPath,

    [parameter()]
    [string]$Backup_Path_Local = $global:backupTempPathLocal
    )
    BEGIN {
        $Backup_FilePath = $Backup_Path + "\$Database.bak"
        $Backup_Path_UNC = PPath-ToUNC -Server $serverInstance -PPath $Backup_FilePath
    }
    PROCESS
    {
        try {
            if ((Test-Path -Path $Backup_Path_Local) -eq $false) {
                Output-Message "Creating temporary directory $Backup_Path_Local"
                New-Item -Path $Backup_Path_Local -ItemType Directory -Force -ErrorAction SilentlyContinue > $null
            }
            Output-Message "Copying file $Backup_FilePath from $serverInstance to temporary path $Backup_Path_Local"
            Copy-Item -Path $Backup_Path_UNC -Destination $Backup_Path_Local -Force -ErrorAction Stop > $null
            return $true
        }
        catch [Exception] {
            return $_
        }
    }
    END {}
}

function ZIP-BackupFile
{
    [cmdletBinding()]
    param(

    [Parameter(Position=1)]
    [String]$InputFile,

    [Parameter(Position=2)]
    [String]$OutputFile,

    [Parameter(Position=3)]
    [String]$ResultFilePath = "E:\zipResultTemp.txt",

    [Parameter(Position=4)]
    [String]$Message = "Zipping up backup file...",

    [Parameter(Position=5)]
    [String]$7zPath = $global:7zPath
    )
    BEGIN{ $Success = $False }
    PROCESS
    {
        try
        {
            $Archive = "$scriptpath\Temp\ZIP\" + $ZIPFile
            $filePath = "$scriptpath\Temp\EAR\" + $EARFile
            $processID = (Start-Process cmd.exe -argumentlist "/c $7zpath a -tzip `"$OutputFile`" `"$InputFile`" -y 1> `"$ResultFilePath`" 2>&1" -PassThru).Id
            Output-Message $Message
            sleep -s 1
            do{
                $ResultFile = Get-Content $ResultFilePath -ErrorAction SilentlyContinue
                Write-Host -NoNewline "#"
                sleep -m 50
            }while ($processID -ne $null -and $ResultFile -eq $null)
            Write-Host ""
            Output-Message "ZIP complete!"
            # Cleanup
            if (Test-Path $ResultFilePath ) {
                try {
                    Remove-Item -Path $ResultFilePath -Force -ErrorAction SilentlyContinue > $null
                }
                catch [Exception] {
                    Output-Message -Silent "[WARNING]: Unable to remove temporary file $ResultFilePath"
                }
            }
            foreach ($line in $ResultFile) {
                if ($line -like "*Everything is OK*")
                {
                    $Success = $true
                }
            }
            if ($Success -eq $true) {
                return $true
            }
            else
            {
                return ("`n" + $ResultFile)
            } 
        }
        catch [Exception] {
            return "$_"
        }
    }
    END{}
}

function FTP-BackupFiles
{
    [CmdletBinding()]
    param(
    [parameter()]
    [string]$ftpserver = $global:ftpserver,

    [parameter()]
    [string]$ftpusername = $global:username,

    [parameter()]
    [string]$ftppassword = $global:password,

    [parameter()]
    [string]$localpath = $global:backupTempZIPPathLocal + '\*',

    [parameter()]
    [string]$remotepath = $global:ftppath

    )
    BEGIN {
        $everything_ok = $True
    }
    PROCESS
    {
        Output-Message "Transferring output through FTP"
        Output-Message "Loading WinSCP Assembliesand"
        try {
            # Load WinSCP .NET assembly
            # Use "winscp.dll" for the releases before the latest beta version.
            [Reflection.Assembly]::LoadFrom("file:////$scriptpath\Res\WinSCP\WinSCP.dll") | Out-Null
            $sessionOptions = New-Object WinSCP.SessionOptions
            $sessionOptions.Protocol = [WinSCP.Protocol]::ftp
            $sessionOptions.HostName = $ftpserver
            $sessionOptions.UserName = $ftpusername
            $sessionOptions.Password = $ftppassword
            $session = New-Object WinSCP.Session
            $session.ExecutablePath = "$scriptpath\Res\WinSCP\WinSCP.exe"
        }
        catch [Exception] {
            Output-Message "Error encountered while setting up WinSCP assembly"
            return "Unable to transfer SQL backup files"
            }
        Output-Message "Initiating FTP session"
        if ($everything_ok -eq $True) {
            try {
                # Connect to FTP Server
                $session.Open($sessionOptions)
                Output-Message "Session initiated"
                # Upload files
                $transferOptions = New-Object WinSCP.TransferOptions
                $transferOptions.TransferMode = [WinSCP.TransferMode]::Automatic

                Output-Message "Transferring output through FTP session"
                $transferResult = $session.PutFiles($localpath,"$remotepath/", $False, $transferOptions)
                # Throw on any error
                $transferResult.Check()
                # Print results
                foreach ($transfer in $transferResult.Transfers) {
				    $fileupload = $transfer.FileName
				    Output-Message "Upload of $fileupload succeeded"
                }
                Output-Message "Output has been successfully transferred to $ftpserver"
            }
            catch [exception] {
                Output-Message "Unable to transfer $localpath to $remotepath"
                return $_
            }
            finally {
                    $session.Dispose()
            }
            if ($transferResult.IsSuccess -eq $True) {
                return $True
            }
        }
    }
    END{}
}

function Clean-TempFiles
{
    param (
    [parameter()]
    [string]$serverInstance,

    [parameter()]
    [string]$Backup_Path = $global:backupTempPath,
    
    [parameter()]
    [string]$Backup_Path_Local = $global:backupTempPathLocal)

    try {
        $Backup_Path_UNC = PPath-ToUNC -Server $serverInstance -PPath $Backup_Path
        if (Test-Path $Backup_Path_UNC) {
            Remove-Item -Path $Backup_Path_UNC -Force -Recurse -ErrorAction Stop > $null
        }
        if (Test-Path $Backup_Path_Local) {
            Remove-Item -Path $Backup_Path_Local -Force -Recurse -ErrorAction Stop > $null
        }
        if (Test-Path $global:EMAILLOGPATH) {
            Remove-Item -Path $global:EMAILLOGPATH -Force -Recurse -ErrorAction Stop > $null
        }
        if (Test-Path $global:backupTempZIPPathLocal) {
            Remove-Item -Path $global:backupTempZIPPathLocal -Force -Recurse -ErrorAction Stop > $null
        }
        return $true
    }
    catch [Exception] {
        return $_
    }
}

function Clean-OldLogs
{
    [CmdletBinding()]
    param(

    [Parameter()]
    [string]$History = $global:logHistory
    )
    try {
        $logsList = get-childitem -Path "$scriptPath\Log" | Sort-Object -Property 'LastWriteTime' -Descending | Select-Object -ExpandProperty Name
        if ($logsList.Count -ge $History) {
            $logCounter = 1
            foreach ($log in $logsList)
            {
                if ($logCounter -gt $History)
                {
                    Remove-Item -Path "$scriptPath\Log\$log" -Force > $null
                }
                $logCounter++
            }
        }
        return $true
    }
    catch [Exception] {
        return $_
    }
}

function Email-Result
{
    [CmdletBinding()]
    param(

    [Parameter(Position=1)]
    [string[]]$To = $global:emailto,

    [Parameter(Position=2)]
    [string]$CC = $global:emailcc,

    [Parameter(Position=3)]
    [string[]]$From = $global:emailfrom,

    [Parameter()]
    [string]$SMTPServer = $global:smtpserver,

    [Parameter()]
    [string]$SMTPPort = $global:smtpport,
       
    [Parameter()]
    [string]$Subject = $global:emailsubject,

    [Parameter()]
    [string]$Body = $global:emailbody,

    [Parameter()]
    [switch]$Error
    )
    BEGIN
    {
    }
    PROCESS
    {
        Output-Message "Creating email message"
        $SMTPmessage = New-Object Net.Mail.MailMessage($From,$To)
        foreach ($address in $CC.Split(','))
        {
            $SMTPmessage.CC.Add($address)    
        }
        $SMTPmessage.Subject = $Subject -replace '<server_name>',$global:serverInstance
        $SMTPmessage.IsBodyHtml = $false
        if ($PSBoundParameters.ContainsKey('Error'))
        {
            $SMTPmessage.Priority = [System.Net.Mail.MailPriority]::High
            $SMTPmessage.Subject = $global:emailerrorsubject -replace '<server_name>',$global:serverInstance
        }
        # Compose email body
        $EmailBody = $Body
        # Filter email text
        $datedisplay = date -F {g}
        $EmailBody = $EmailBody -replace '<server_name>',$global:serverInstance
        $EmailBody = $EmailBody -replace '<exec_date>',$datedisplay
        $EmailBody = $EmailBody -replace '<exec_duration>',$global:duration
        # Output Database List
        if ($global:databases.Count -gt 1) {
            $EmailBody += "`nDATABASES:`n"
            $i = 1
            foreach ($database in $global:databases) {
                $EmailBody += "($i) $database`n"
                $i++
            }
            if ($global:filteredDBs -ne $null) {
                foreach ($database in $global:filteredDBs) {
                    $EmailBody += "($i) $database (Excluded from backup)`n"
                    $i++
                }
            }
        }
        else {
            $EmailBody += "Database: $global:databases`n"
        }
        if ($PSBoundParameters.ContainsKey('Error'))
        {
            $EmailBody += "`n[RESULT]: ERROR"
            $EmailBody += "`n`n------ Refer to attached log file for details ------"
        }
        else
        {
            $EmailBody += "`n[RESULT]: SUCCESS"
            $EmailBody += "`n`n------ Refer to attached log file for details ------"
        }

        ## End email body composition
        $SMTPmessage.Body = $EmailBody
        Output-Message "Email message created"
        Copy-Item -Path $global:LOGPATH -Destination $global:EMAILLOGPATH -Force -ErrorAction SilentlyContinue > $null
        $Attached_log = New-Object System.Net.Mail.Attachment($global:EMAILLOGPATH)
        $SMTPmessage.attachments.Add($Attached_log)
        $SMTPClient = New-Object Net.Mail.SmtpClient($SMTPServer,$SMTPPort)
        try
        {
            Output-Message "Sending email message..."
            $SMTPClient.Send($SMTPmessage)
            Output-Message "Email sent!"
            return $true
        }
        catch [Exception]
        {
            Output-Message "Unable to send email" -Red
            return $_
        }
    }
    END
    {
        
        $Attached_log.Dispose()
        $SMTPmessage.Dispose()
    }
}

# Supporting Functions
function Output-Message
{
    [CmdletBinding()]
    param(

    [Parameter(Position=1)]
    [String]$Message,

    [Parameter(Position=2)]
    [String]$Path = $global:LOGPATH,

    [Parameter(Position=3)]
    [switch]$Green,

    [Parameter(Position=4)]
    [switch]$Yellow,

    [Parameter(Position=5)]
    [switch]$Red,

    [Parameter(Position=6)]
    [switch]$Silent,

    [Parameter(Position=7)]
    [switch]$NoLog
    )

    if (($PSBoundParameters.ContainsKey('Silent')) -ne $true)
    {
        if ($PSBoundParameters.ContainsKey('Green')) { Write-Host -ForegroundColor Green $Message }
        elseif ($PSBoundParameters.ContainsKey('Yellow')) { Write-Host -ForegroundColor Yellow $Message }
        elseif ($PSBoundParameters.ContainsKey('Red')) { Write-Host -ForegroundColor Red $Message }
        else { Write-Host $Message }
    }
    if (($PSBoundParameters.ContainsKey('NoLog')) -ne $true) {
        $datedisplay = date -F {MM/dd/yyy hh:mm:ss:}
        [IO.File]::AppendAllText($Path,"$datedisplay $Message`r`n")
    }
}

function PPath-ToUNC
{
    [CmdletBinding()]
    param(

    [Parameter(Position=1)]
    [String]$Server,

    [Parameter(Position=2)]
    [String]$PPath

    )
    BEGIN
    {
    }
    PROCESS
    {
        $driveLetter = $PPath[0]
        $UNC = '\\' + $server + '\' + $driveLetter + '$' + ($PPath -replace "$driveLetter`:",'') 
        $UNC
    }
    END
    {
    }
}

#-- Definition functions to decript encrypted FTP credetial
function Get-DecryptedFTPUser {
    [CmdletBinding()]
    param(
    [parameter()]
    $DecryptKey,
    
    [parameter()]
    [string]$Path
    )

    $ftpcred = ((Get-Content -Delimiter "," "$Path") -replace ',','')
    $ftpuser =  Decrypt-SecureString (ConvertTo-SecureString $ftpcred[0] -Key $DecryptKey)
    $ftpuser
}

function Get-DecryptedFTPPassword {
    [CmdletBinding()]
    param(
    [parameter()]
    $DecryptKey,
    
    [parameter()]
    [string]$Path
    )

    $ftpcred = ((Get-Content -Delimiter "," "$Path") -replace ',','')
    $ftppassword = Decrypt-SecureString (ConvertTo-SecureString $ftpcred[1] -Key $DecryptKey)
    $ftppassword
}

function Decrypt-SecureString {
    param(
    [Parameter(ValueFromPipeline=$true,Mandatory=$true,Position=0)]
    [System.Security.SecureString]$sstr
    )

    $marshal = [System.Runtime.InteropServices.Marshal]
    $ptr = $marshal::SecureStringToBSTR( $sstr )
    $str = $marshal::PtrToStringBSTR( $ptr )
    $marshal::ZeroFreeBSTR( $ptr )
    $str
}

function Get-DecryptedPKPassword
{
    [CmdletBinding()]
    param(
    [parameter()]
    $DecryptKey,
    
    [parameter()]
    [string]$Path
    )

    $pkcontent = Get-Content $Path
    $pkpassword = Decrypt-SecureString (ConvertTo-SecureString $pkcontent -Key $DecryptKey)
    $pkpassword
}
##

# Execute
Main-Program