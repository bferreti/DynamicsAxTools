$Conn = New-Object System.Data.SqlClient.SQLConnection
$Conn.ConnectionString = "Server=UDBSQCR3-MAX\MAX;Database=tempdb;Integrated Security=True;Connect Timeout=5"
$Conn.Open()
$Query = Get-Content C:\Users\591607466\Desktop\Bruno_Files\ConPeek.sql | Out-String
$Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Conn)
$Cmd.ExecuteScalar()

$Query = Get-Content C:\Users\591607466\Desktop\Bruno_Files\ConSize.sql | Out-String
$Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Conn)
$Cmd.ExecuteScalar()

$Conn.Close()