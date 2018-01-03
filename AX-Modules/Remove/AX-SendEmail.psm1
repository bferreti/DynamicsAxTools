function Send-Email
{
param (
    [String]$Subject,
    [String]$Body,
    [String]$Attachment,
    [String]$EmailProfile,
    [String]$Guid
)
    $Conn = New-Object System.Data.SqlClient.SQLConnection(Get-ConnectionString)
    $Query = "SELECT * FROM [AXTools_EmailProfile] AS A
                JOIN [AXTools_UserAccount] AS B ON A.UserID = B.ID
                WHERE A.ID = '$EmailProfile'"
    $Adapter = New-Object System.Data.SqlClient.SqlDataAdapter($Query, $Conn)
    $Table = New-Object System.Data.DataSet
    $Adapter.Fill($Table) | Out-Null

    if (![string]::IsNullOrEmpty($Table.Tables))
    {
        $SMTPServer = $Table.Tables.SMTPServer
        $SMTPPort = $Table.Tables.SMTPPort
        $SMTPUserName = $Table.Tables.UserName
        $SMTPPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($($Table.Tables.Password | ConvertTo-SecureString)))
        $SMTPSSL = $Table.Tables.SMTPSSL
        $SMTPFrom = $Table.Tables.From
        $SMTPTo = $Table.Tables.To
        $SMTPCC = $Table.Tables.CC
        $SMTPBCC = $Table.Tables.BCC
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
    if($SMTPSSL -eq 1) {
        $SMTPClient.EnableSsl = $true
        [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { return $true }
    }

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

    SQL-BulkInsert AXTools_EmailLogs @($SMTPMessage | Select @{n='Sent';e={$Sent}}, 
                                        @{n='EmailProfile';e={$EmailProfile}},
                                        @{n='Subject';e={$Subject}},
                                        @{n='Body';e={$Body}},
                                        @{n='Attachment';e={$Attachment}}, 
                                        @{n='Log';e={$Log}},
                                        @{n='GUID';e={$Guid}})

    if($Attachment) {
        $AttachmentFile.Dispose()
    }  
}