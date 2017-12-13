function Send-Email
{
param (
    [String]$Subject,
    [String]$Body,
    [String]$Attachment,
    [String]$EmailProfile,
    [String]$GUID
)
    $Conn = New-Object System.Data.SqlClient.SQLConnection(Get-ConnectionString)
    $Query = "SELECT * FROM [AXTools_EmailProfile] AS Email
                JOIN [AXTools_AccountProfile] AS Acct ON Email.ConnectionID = Acct.ID
                WHERE Email.PROFILEID = '$EmailProfile'"
    $Adapter = New-Object System.Data.SqlClient.SqlDataAdapter($Query, $Conn)
    $Table = New-Object System.Data.DataSet
    $Adapter.Fill($Table) | Out-Null

    if (![string]::IsNullOrEmpty($Table.Tables))
    {
        $SMTPServer = $(Read-EncryptedString -InputString $Table.Tables.ConnectionInfo.Split(',')[0] -DTKey "$((Get-WMIObject Win32_Bios).PSComputerName)-$((Get-WMIObject Win32_Bios).SerialNumber)")
        $SMTPPort = $(Read-EncryptedString -InputString $Table.Tables.ConnectionInfo.Split(',')[1] -DTKey "$((Get-WMIObject Win32_Bios).PSComputerName)-$((Get-WMIObject Win32_Bios).SerialNumber)")
        $SMTPUserName = $(Read-EncryptedString -InputString $Table.Tables.Data.Split(',')[0] -DTKey "$((Get-WMIObject Win32_Bios).PSComputerName)-$((Get-WMIObject Win32_Bios).SerialNumber)")
        $SMTPPassword = $(Read-EncryptedString -InputString $Table.Tables.Data.Split(',')[1] -DTKey "$((Get-WMIObject Win32_Bios).PSComputerName)-$((Get-WMIObject Win32_Bios).SerialNumber)")
        $SMTPSSL = $(Read-EncryptedString -InputString $Table.Tables.ConnectionInfo.Split(',')[2] -DTKey "$((Get-WMIObject Win32_Bios).PSComputerName)-$((Get-WMIObject Win32_Bios).SerialNumber)")
        $SMTPFrom = $($Table.Tables.From)
        $SMTPTo = $($Table.Tables.To)
        $SMTPCC = $($Table.Tables.CC)
        $SMTPBCC = $($Table.Tables.BCC)
        $Table.Dispose()
    }
    else {
        break
    }

    #Message Parameters
    $SMTPMessage = New-Object System.Net.Mail.MailMessage
    $SMTPMessage.From = $SMTPFrom
    if(-not [System.DBNull]::Value.Equals($SMTPTo)) {$SMTPTo.Split(';') | % { $SMTPMessage.To.Add($_.Trim()) }} else { break }
    if(-not [System.DBNull]::Value.Equals($SMTPCC)) {$SMTPCC.Split(';') | % { $SMTPMessage.CC.Add($_.Trim()) }}
    if(-not [System.DBNull]::Value.Equals($SMTPBCC)) {$SMTPBCC.Split(';') | % { $SMTPMessage.Bcc.Add($_.Trim()) }}
    $SMTPMessage.Subject = $Subject
    $SMTPMessage.IsBodyHtml = $true
    $SMTPMessage.Body = $Body

    #Attachemnts
    if($Attachment) {
        $AttachmentFile = New-Object System.Net.Mail.Attachment($Attachment)
        $SMTPMessage.Attachments.Add($AttachmentFile)
    }

    #Create Message
    $SMTPClient = New-Object System.Net.Mail.SmtpClient($SMTPServer,$SMTPPort)
    $SMTPClient.Credentials = New-Object System.Net.NetworkCredential($SMTPUserName,$SMTPPassword)

    #Send Email
    if($SMTPSSL -like 'True') {
        $SMTPClient.EnableSsl = $true
        [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { return $true }
    }
    
    #$SMTPClient.Send($SMTPMessage)

    try
    {
        $SMTPClient.Send($SMTPMessage)
        $Sent = '1'
        $Log = ''
        SQL-ExecUpdate "UPDATE AXMonitor_ExecutionLog SET EMAIL = '1' WHERE GUID = '$Guid'"
    }
    catch
    {
        $Sent = '0'
        $Log = $_.Exception
    }

    SQL-BulkInsert AXTools_EmailLog @($SMTPMessage | Select @{n='Sent';e={$Sent}}, 
                                        @{n='EmailProfile';e={$EmailProfile}},
                                        @{n='Subject';e={$Subject}},
                                        @{n='Body';e={$Body}},
                                        @{n='Attachment';e={$Attachment}}, 
                                        @{n='Log';e={$Log}},
                                        @{n='GUID';e={$GUID}})

    if($Attachment) {
        $AttachmentFile.Dispose()
    }  
}




