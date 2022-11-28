using namespace System.Data
using namespace System.Data.SqlClient
using namespace System.Data.SQLite
using namespace System.IO

class DbProvider {

    [SQLiteConnection] $LocalConnection
    [SqlConnection] $RemoteConnection 
    

    DbProvider ([string]$ConnectionStirng){
        $this.LocalConnection = [SQLiteConnection]::new()
        $this.LocalConnection.ConnectionString = "Data source=$PSScriptRoot/database.db;Version=3"
        $this.LocalConnection.Open()

        $this.RemoteConnection = [SqlConnection]::new()
        $this.RemoteConnection.ConnectionString = $ConnectionStirng
    }

    [string[]] GetClientIdArray () {
        $res = @()
        $c = $this.LocalConnection.CreateCommand()
        $c.CommandText = "select ID from CLIENT"
        $dr = $c.ExecuteReader()
        while ($dr.Read()){
            $res +=, $dr[0]
        }
        $dr.Close()
        return $res
    }
}



