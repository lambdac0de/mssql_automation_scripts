# ----- Create-Credential -----
# This script will create an encrypted credential file based on the specified key below
# Encryption method uses Microsoft's Data Protection API, which is a symmetric key encryption
# ----------------------------------------------------------------------------

# Specify encryption key here
# This is an example key, you can create your own by simply randomizing and replacing the numbers below
$Key = (1,4,2,3,56,34,254,222,1,1,2,23,42,54,33,233,1,34,2,7,6,5,35,43)
##

#Function to create encrypted credential
function Create-EncryptedCredential
{
    param(
    [Parameter(Mandatory=$true,Position=1)]
    [string]$username,

    [Parameter(Mandatory=$true,Position=2)]
    [string]$password,

    [Parameter(Mandatory=$true,Position=3)]
    [string]$outputfile
    )

    $usernameenc = ConvertTo-SecureString $username -AsPlainText -Force
    $uc = ConvertFrom-SecureString -SecureString $usernameenc -Key $Key

    $passwordenc = ConvertTo-SecureString $password -AsPlainText -Force
    $pc = ConvertFrom-SecureString -SecureString $passwordenc -Key $Key

    echo "$uc,$pc" | Out-File $outputfile
}
##

#Prompt for credentials to encrypt
$username = Read-Host 'Enter username'
$passinput = Read-Host 'Enter password' -AsSecureString
$password = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($passinput))
##

#Create the credential file!
Create-EncryptedCredential -username $username -password $password -outputfile "ftpcred.enc"
##

#Show confirmation message
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") > $null
[System.Windows.Forms.MessageBox]::Show("Credential file 'ftpcred.enc' has been created" , "Success") > $null
##