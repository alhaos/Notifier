using namespace System.Data
using namespace System.Data.SqlClient
using namespace System.Data.SQLite
using namespace System.IO

class DbProvider {

    [SQLiteConnection] $LocalConnection
    [SqlConnection] $RemoteConnection
    [string]$ConnectionString

    DbProvider ([string]$ConnectionString) {
        $this.ConnectionString = $ConnectionString
        $this.LocalConnection = [SQLiteConnection]::new()
        $this.LocalConnection.ConnectionString = "Data source=$PSScriptRoot/database.db;Version=3"
        $this.LocalConnection.Open()

        $this.RemoteConnection = [SqlConnection]::new()
        $this.RemoteConnection.ConnectionString = $ConnectionString
    }

    [string] GetClientIdArray () {
        $res = @()
        $c = $this.LocalConnection.CreateCommand()
        $c.CommandText = "select ID from CLIENT"
        $dr = $c.ExecuteReader()
        while ($dr.Read()) {
            $res += , "'{0}'" -f $dr[0]
        }
        $dr.Close()
        return $res -join ","
    }
    
    #LoadData load data to RAW_DATA table
    LoadData () {
        Write-LogInfo "LoadData start"
        $c = [SqlConnection]::new()
        $c.ConnectionString = $this.ConnectionString
        $cm = $c.CreateCommand()
        $cm.CommandText = @"
        select
        l.Accession                                     [Accession]
        , TRY_CONVERT(date, l.[Final Report Date])      [Final Report Date]
        , RTRIM(l.[Patient First Name])                 [First Name]
        , RTRIM(l.[Patient Last Name])                  [Last Name]
        , l.MI                                          [Middle Name]
        , l.DOB                                         [DOB]
        , l.[Client ID]                                 [Client ID]
        , l.[Client Name]                               [Client Name]
        , l.[Phys  Name]                                [Phys Name]
        , l.[Test Code]                                 [Test Code]
        , l.[Test Name]                                 [Test Name]
        , RTRIM(LTRIM(l.Result))                        [Test Result]
        , l.[Patient Address]                           [Patient Address]
        , l.[Patient City]                              [Patient City]
        , l.[Patient State]                             [Patient State]
        , l.[Patient Zip]                               [Patient Zip]
        , l.[Patient Phone]                             [Patient Phone]
    from logtest l
    where CAST(l.[Final Report Date] as date) > getdate() - 2
        and l.[Test Code] in ('950Z','960Z','Z642','Z801','Z620','620Z')
        and l.[Final Report Date] is not null
        and LTRIM(l.[Final Report Date]) != ''
        and LTRIM(l.[DOB]) != ''
        and rtrim(l.result) not in ('ON', 'ONGOING')
        and l.[Client ID] in (#ClinetIdTag#)
    order by [Final Report Date], [Client ID]
"@ -replace '#ClinetIdTag#', $this.GetClientIdArray()
        $c.Open()
        $cmdDelete = $this.LocalConnection.CreateCommand()
        $cmdDelete.CommandText = "delete from RAW_DATA;"
        $cmdDelete.ExecuteNonQuery()
        $cmdInsert = $this.LocalConnection.CreateCommand()
        $dr = $cm.ExecuteReader()
        $i = 0
        while ($dr.Read()) {
            $i++
            $cmdInsert.CommandText = @'
insert into RAW_DATA (
    ACCESSION,"FINAL REPORT DATE","FIRST NAME","LAST NAME","MIDDLE NAME",DOB,"CLIENT ID","CLIENT NAME","PHYS NAME","TEST CODE","TEST NAME","TEST RESULT","PATIENT ADDRESS","PATIENT CITY","PATIENT STATE","PATIENT ZIP","PATIENT PHONE"
)
values (
'{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}','{15}','{16}'
);
'@ -f $dr[0], $dr[1].ToString("yyyy-MM-dd HH:mm:ss"), $dr[2].Replace("'", '`'), $dr[3].Replace("'", '`'), $dr[4], $dr[5], $dr[6], $dr[7], $dr[8], $dr[9], $dr[10], $this.InterpretResult($dr[11]), $dr[12], $dr[13], $dr[14], $dr[15], $dr[16]
            Write-Debug $cmdInsert.CommandText
            $cmdInsert.ExecuteNonQuery()
        }
        $c.Close()
        Write-LogInfo "LoadData fatched $i records"
        Write-LogInfo "LoadData end"
    }

    [string] InterpretResult ([string] $Result) {
        $a = switch ($Result) {
            'ND' {
                'Not Detected' 
            }
            'ON' {
                'Ongoing' 
            }
            'INV' {
                'Invalid' 
            }
            'D' {
                'Detected' 
            }
            'IN' {
                'Indeterminate' 
            }
            Default { "Error" }
        }
        return $a
    }
}