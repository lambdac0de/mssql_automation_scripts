# mssql_automation_scripts
This repository contains scritps to automate common tasks and workflows associated with MS SQL servers<br><br>
<b>WARNING:</b> This repository mostly contains a set of automation scripts I developed in 2015. The growth of PowerShell has been rapid since then, so there are most likely better, more efficient, ways of doing the same implementations.
## mssql_backup_automation
#### What is this?
This is an automation script to perform a trial yet common task: backup databases, compress them, and upload to an FTP server. This provides a very simple, yet effective, backup and DR solution for MS SQL databases without a need for expensive enterprise solutions.
#### Prerequisites
1. 7zip binaries (included)
2. sqlps module (included)
3. WinSCP binaries (included)
4. PowerShell 2.0
#### How this works
1. Lists all user databases from a given SQL Server instance
2. Backups all user databases
3. Verifies all backups as valid
4. Compresses all backups (using 7zip)
5. Uploads compressed backups via FTP (using WinSCP)
6. Sends email notification to intended recipients about task completion
#### Usage
1. Create your FTP credential (See `Create-FTPcred.ps1`) and make sure the account has proper write permissions
2. Ensure `config.ini` is properly populated with correct config values (SQL Server, FTP, and Email settings)
3. Review `.\Res` folder and ensure all prerequisites are there
4. Make a `Log` folder in the program root, or change log path in variable `$global:LOGPATH`
5. Run the program, or schedule as needed<br><br>
<b>Note</b> It is assumed that the user (service) account runing the script has appropriate permission on the target MSSQL server. The script will use integrated authentication using the impersonated user context.
