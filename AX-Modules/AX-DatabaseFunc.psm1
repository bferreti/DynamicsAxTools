$ParamDBServer = 'UDBSQCR3-MAX\MAX' #Change SQL Server Name. Server\Instance, (local)
$ParamDBName = 'DynamicsAXTools' #Change DB Name (if you have changed during creation)

function Get-ConnectionString {
    return "Server=$ParamDBServer;Database=$ParamDBName;Integrated Security=True;Connect Timeout=5"
}