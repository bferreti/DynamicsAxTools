[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.SMO') | Out-Null

function Get-ConnectionString {
    return "Server=$Computer;Database=$MessageDB;Integrated Security=True;Connect Timeout=5"
}

$FileDateTime = Get-Date -f yyyyMMdd-HHmm

$Servers = Get-Content C:\Users\chiconatoferretib\Desktop\Bruno\AAFES-Powershell\Channels.txt
#$Computer = 'S106920500'

foreach($Computer in $Servers) {
    Write-Host "Connectiing to $Computer." -fore Green
    try {
    $Server = New-Object ('Microsoft.SqlServer.Management.Smo.Server') $Computer  #| Out-Null
    $MessageDBs = $Server.Databases.Name | Where {$_ -match "CHMSGDB"}
    }
    catch {
        continue
    }
    if($MessageDBs) {
        foreach($MessageDB in $MessageDBs) {
            Write-Host "--> DB $MessageDB - " -fore Yellow -NoNewline
            $Conn = New-Object System.Data.SqlClient.SQLConnection(Get-ConnectionString)
            $Query = " --Async Client
                            SELECT TOP 1000 B.DOWNLOADSESSIONID, MAX(A.STATUS) AS STATUS, 
				                            CASE MAX(A.STATUS)
		                                            WHEN 0 THEN 'Started'
						                            WHEN 1 THEN 'Available'
						                            WHEN 2 THEN 'Requested'
						                            WHEN 3 THEN 'Downloaded'
						                            WHEN 4 THEN 'Applied'
						                            WHEN 5 THEN 'Canceled'
		                                            WHEN 6 THEN 'Create Failed'
		                                            WHEN 7 THEN 'Download Failed'
					                                WHEN 8 THEN 'Apply Failed'
						                            WHEN 9 THEN 'No Data'
	                                                END AS STATUS_DESC,
			                            MAX(B.DATECREATED) AS LASTCHANGE, COUNT(*) AS CNT, A.JOBID, @@SERVERNAME AS ServerName, DB_NAME() as DBName
                            FROM DOWNLOADSESSION A
                            JOIN DOWNLOADSESSIONLOG B ON A.ID = b.DOWNLOADSESSIONID
                            WHERE B.DATECREATED >= '$((Get-Date).AddDays(-1).Date)'
                            GROUP BY B.DOWNLOADSESSIONID, A.JOBID
                            HAVING MAX(A.STATUS) <> 4
                            ORDER BY 4 DESC"

            $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Conn)
            $Adapter = New-Object System.Data.SqlClient.SqlDataAdapter
            $Adapter.SelectCommand = $Cmd
            $CDXJobs = New-Object System.Data.DataSet
            $Adapter.Fill($CDXJobs)
            $Conn.Close()

            $CDXJobs.Tables[0] | Export-Csv C:\Users\chiconatoferretib\Desktop\Bruno\CdxJobs_$($FileDateTime).csv -NoTypeInformation -Append
        }
    }
}