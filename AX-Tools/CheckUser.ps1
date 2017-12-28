Import-Module C:\Users\Administrator\Documents\GitHub\DynamicsAxTools\AX-Modules\AX-StringCrypto.psm1
Remove-Module -Name AX-StringCrypto

$SMTPUser = 'ps.monitoring.pfe@gmail.com'
$SMTPPass = 'igyxkgbbvtmgoofa'

$SMTPServer = 'smtp.gmail.com'
$SMTPPort = '587'
$SMTPSSL = $true
$PasswordToEncrypt = "$((Get-WMIObject Win32_Bios).PSComputerName)-$((Get-WMIObject Win32_Bios).SerialNumber)"


$SMTPUserEnc = Write-EncryptedString -Inputstring $SMTPUser -DTKey $PasswordToEncrypt
$SMTPPassEnc = Write-EncryptedString -Inputstring $SMTPPass -DTKey $PasswordToEncrypt

$SMTPServerEnc = Write-EncryptedString -Inputstring $SMTPServer -DTKey $PasswordToEncrypt
$SMTPPortEnc = Write-EncryptedString -Inputstring $SMTPPort -DTKey $PasswordToEncrypt
$SMTPSSLEnc = Write-EncryptedString -Inputstring $SMTPSSL -DTKey $PasswordToEncrypt

$SMTPConn = "$SMTPServerEnc,$SMTPPortEnc,$SMTPSSLEnc"

$SMTPConn = "$SMTPUserEnc,$SMTPPassEnc"


#Read-EncryptedString -InputString $SMTPConn -DTKey $PasswordToEncrypt


$Conn = New-Object System.Data.SqlClient.SQLConnection(Get-ConnectionString)
$Query = "INSERT INTO [dbo].[AX_EmailProfile] ([PROFILEID],[CONNECTIONINFO],[TO],[CC])
            VALUES ('MFRM_uat','$SMTPConn','bferreti@microsoft.com','')"

$Conn.Open()
$Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Conn)
$Cmd.ExecuteNonQuery()
$Conn.Close()

$Conn = New-Object System.Data.SqlClient.SQLConnection(Get-ConnectionString)
$Query = "UPDATE [dbo].[AX_EmailProfileNew] SET [CONNECTIONINFO] = '$SMTPConn'
            WHERE PROFILEID = 'MFRMNew'"

$Conn.Open()
$Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Conn)
$Cmd.ExecuteNonQuery()
$Conn.Close()


$Conn = New-Object System.Data.SqlClient.SQLConnection(Get-ConnectionString)
$Query = "INSERT INTO [dbo].[AX_ConnectionProfile] ([ID],[Data])
            VALUES ('$SMTPUser','$SMTPConn')"

$Conn.Open()
$Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Conn)
$Cmd.ExecuteNonQuery()
$Conn.Close()


#(Get-WMIObject Win32_Bios) | FL *

#(Get-WMIObject Win32_Bios).PSComputerName