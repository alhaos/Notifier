using namespace System.Data
using namespace System.Data.SQLite
using namespace System.Data.SqlClient
using namespace System.IO

Set-StrictMode -Version 'Latest'

class Notifier {
    [hashtable]$Config
    [Rep[]]$Reps
    [Client[]]$Clients
    [SQLiteConnection]$LocalConnection = [SQLiteConnection]::new()
    [SqlConnection]$RemoteConnection = [SqlConnection]::new()
    
    Notifier([hashtable]$Config) {
        $this.Config = $Config
        $this.LocalConnection.ConnectionString = "Data source=$PSScriptRoot/database.db;Version=3"        
        $this.LocalConnection.Open()
        $this.RemoteConnection.ConnectionString = $this.Config.ConnectionString
    }

    LoadData () {
        Write-LogInfo "LoadData start"
        $c = [SqlConnection]::new()
        $c.ConnectionString = $this.Config.ConnectionString
        $cm = $c.CreateCommand()
        $cm.CommandText = @"
            select
            l.Accession                                     [Accession]
            , l.[Final Report Date]                         [Final Report Date]
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
            left join LTEDB.dbo.COVID_MAILOUT mo on l.Accession = mo.Accession and l.[Test Code] = mo.TestCode
        where CAST(l.[Final Report Date] as date) > getdate() - 2
            and l.[Test Code] in ('950Z', '960Z')
            and l.[Final Report Date] is not null
            and LTRIM(l.[Final Report Date]) != ''
            and LTRIM(l.[DOB]) != ''
            and mo.Accession is NULL
            and rtrim(l.result) not in ('ON', 'ND')
            and l.[Client ID] in (#ClinetIdTag#)
        order by [Final Report Date], [Client ID]
"@ -replace '#ClinetIdTag#', $this.GetClientList()
            
        $c.Open()
        $cmdDelete = $this.LocalConnection.CreateCommand()
        $cmdDelete.CommandText = "delete from RAW_DATA;"
        $null = $cmdDelete.ExecuteNonQuery()
        $cmdInsert = $this.LocalConnection.CreateCommand()
        $dr = $cm.ExecuteReader()
        $i = 0
        while ($dr.Read()) {
            $i++
            $commandTest = @'
    insert into RAW_DATA (
        ACCESSION,"FINAL REPORT DATE","FIRST NAME","LAST NAME","MIDDLE NAME",DOB,"CLIENT ID","CLIENT NAME","PHYS NAME","TEST CODE","TEST NAME","TEST RESULT","PATIENT ADDRESS","PATIENT CITY","PATIENT STATE","PATIENT ZIP","PATIENT PHONE"
    )
    values (
    '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}','{15}','{16}'
    );
'@
            $cmdInsert.CommandText = $commandTest -f (
                $dr[0], $dr[1], $dr[2], $dr[3], $dr[4], $dr[5], $dr[6], $dr[7],
                $dr[8], $dr[9], $dr[10], ($this.InterpretResult($dr[11])),
                $dr[12], $dr[13], $dr[14], $dr[15], $dr[16]
            )
                
            $null = $cmdInsert.ExecuteNonQuery()
        }
        $c.Close()
        $cmdDelete.CommandText = "delete from RAW_DATA where ACCESSION in (select ACCESSION from PROCESSED_ACCESSION);"
        $cmdDelete.ExecuteNonQuery()

        $cmdDelete.CommandText = "delete from PROCESSED_ACCESSION where DT < @DT;"
        $prmDt = [SQLiteParameter]::new("@Dt", [DbType]::String)
        $cmdDelete.Parameters.Add($prmDt)
        $prmDt.Value = "{0:yyyyMMdd HHmmss}" -f [datetime]::Now.AddDays(-100)
        $cmdDelete.ExecuteNonQuery()


        Write-LogInfo "LoadData end"
    }

    FillReps () {
        $cmdRep = $this.LocalConnection.CreateCommand()
        $cmdRep.CommandText = @"
select distinct r.ID, r.NAME from RAW_DATA rd
  join CLIENT c on rd."CLIENT ID" = c.ID
  join REP r on c.REP_ID = r.ID
 where r.name != 'None'
"@
        $cmdRepEmail = $this.LocalConnection.CreateCommand()
        $cmdRepEmail.CommandText = @"
select et.NAME EMAIL_TYPE_NAME,
       re.ADDRESS
  from REP_EMAIL re
  join EMAIL_TYPE et 
    on re.TYPE_ID = et.ID
 where REP_ID = @ID 
"@
        $idParameter = [SQLiteParameter]::new("@ID", [DbType]::String)
        $cmdRepEmail.Parameters.Add($idParameter)

        $cmdTests = $this.LocalConnection.CreateCommand()
        $cmdTests.CommandText = @"
select rd.* from RAW_DATA rd
  join CLIENT c on rd."CLIENT ID" = c.ID
  join REP r on c.REP_ID = r.ID
  join RESULT rt on rt.NAME = rd."TEST RESULT"
  join CLIENT_RESULT cr on c.id = cr.CLIENT_ID and rt.ID = cr.RESULT_ID
 where r.name = @NAME
"@
        $nameParameter = [SQLiteParameter]::new("@NAME", [DbType]::String)
        $cmdTests.Parameters.Add($nameParameter)

        $dr = $cmdRep.ExecuteReader()
        while ($dr.Read()) {
            $rep = [Rep]@{
                Name = $dr['NAME']
            }
            $idParameter.Value = $dr['ID']
            $drEmail = $cmdRepEmail.ExecuteReader()
            while ($drEmail.Read()) {
                $t = $drEmail["EMAIL_TYPE_NAME"]
                $rep.$t += , $drEmail["ADDRESS"]
            }
            $drEmail.Close()

            $nameParameter.Value = $dr['NAME']
            $testDr = $cmdTests.ExecuteReader()
            while ($testDr.Read()) {
                $test = [Test]@{
                    'ACCESSION'         = $testDr['ACCESSION']
                    'Final Report Date' = $testDr['Final Report Date']
                    'First Name'        = $testDr['First Name']
                    'Last Name'         = $testDr['Last Name']
                    'Middle Name'       = $testDr['Middle Name']
                    'DOB'               = $testDr['DOB']
                    'Client ID'         = $testDr['Client ID']
                    'Client Name'       = $testDr['Client Name']
                    'Phys Name'         = $testDr['Phys Name']
                    'Test Code'         = $testDr['Test Code']
                    'Test Name'         = $testDr['Test Name']
                    'Test Result'       = $testDr['Test Result']
                    'Patient Address'   = $testDr['Patient Address']
                    'Patient City'      = $testDr['Patient City']
                    'Patient State'     = $testDr['Patient State']
                    'Patient Zip'       = $testDr['Patient Zip']
                    'Patient Phone'     = $testDr['Patient Phone']
                }
                $rep.TestArray += , $test
            }
            $testDr.Close()

            if ($rep.TestArray.Count) {
                $this.Reps += , $rep
            }
        }
        $dr.Close()
    }

    FillClients () {
        $cmdClient = $this.LocalConnection.CreateCommand()
        $cmdClient.CommandText = 'select distinct "CLIENT ID" from RAW_DATA'

        $cmdClientEmail = $this.LocalConnection.CreateCommand()
        $cmdClientEmail.CommandText = @"
 select et.NAME EMAIL_TYPE_NAME,
        ce.ADDRESS
   from CLIENT_EMAIL ce
   join EMAIL_TYPE et
     on ce.TYPE_ID = et.ID
  where CLIENT_ID = @ID
"@
        $idParameter = [SQLiteParameter]::new("@ID", [DbType]::String)
        $cmdClientEmail.Parameters.Add($idParameter)

        $cmdTests = $this.LocalConnection.CreateCommand()
        $cmdTests.CommandText = @"
select rd.* from RAW_DATA rd
  join CLIENT c on rd."CLIENT ID" = c.ID
  join RESULT rt on rt.NAME = rd."TEST RESULT"
  join CLIENT_RESULT cr on c.id = cr.CLIENT_ID and rt.ID = cr.RESULT_ID
 where rd."CLIENT ID" = @ID
"@
        $cmdTests.Parameters.Add($idParameter)

        $dr = $cmdClient.ExecuteReader()
        while ($dr.Read()) {
            $Client = [Client]@{
                Name = $dr['CLIENT ID']
            }
            $idParameter.Value = $dr["CLIENT ID"]
            $drEmail = $cmdClientEmail.ExecuteReader()
            while ($drEmail.Read()) {
                $t = $drEmail["EMAIL_TYPE_NAME"]
                $Client.$t += , $drEmail["ADDRESS"]
            }
            $drEmail.Close()

            $testDr = $cmdTests.ExecuteReader()
            while ($testDr.Read()) {
                $test = [Test]@{
                    'ACCESSION'         = $testDr['ACCESSION']
                    'Final Report Date' = $testDr['Final Report Date']
                    'First Name'        = $testDr['First Name']
                    'Last Name'         = $testDr['Last Name']
                    'Middle Name'       = $testDr['Middle Name']
                    'DOB'               = $testDr['DOB']
                    'Client ID'         = $testDr['Client ID']
                    'Client Name'       = $testDr['Client Name']
                    'Phys Name'         = $testDr['Phys Name']
                    'Test Code'         = $testDr['Test Code']
                    'Test Name'         = $testDr['Test Name']
                    'Test Result'       = $testDr['Test Result']
                    'Patient Address'   = $testDr['Patient Address']
                    'Patient City'      = $testDr['Patient City']
                    'Patient State'     = $testDr['Patient State']
                    'Patient Zip'       = $testDr['Patient Zip']
                    'Patient Phone'     = $testDr['Patient Phone']
                }
                $Client.TestArray += , $test
            }
            $testDr.Close()

            $this.Clients += , $Client
        }
        $dr.Close()
    }
    
    SendReps () {
        foreach ($rep in $this.Reps) {
            $splat = @{
                To       = $rep.To
                Cc       = $rep.Cc
                Bcc      = $rep.Bcc
                HtmlBody = $rep.GetHtmlBody()
                Subject  = $rep.Name
            }
            $null = Send-AccuMail @splat
            Write-LogInfo ("{0} sent" -f $rep.Name)
        }
    }

    SendClients() {
        $cmd = $this.LocalConnection.CreateCommand()
        $cmd.CommandText = "insert into PROCESSED_ACCESSION (ACCESSION, DT) values (@ACCESSION, @DT)"

        $prmAcc = [SQLiteParameter]::new("@ACCESSION", [DbType]::String)
        $prmDt = [SQLiteParameter]::new("@DT", [DbType]::String)

        $cmd.Parameters.Add($prmAcc)
        $cmd.Parameters.Add($prmDt)

        foreach ($client in $this.Clients) {
            $splat = @{
                To       = $client.To
                Cc       = $client.Cc
                Bcc      = $client.Bcc
                HtmlBody = $client.GetHtmlBody()
                Subject  = $client.Name
            }
            if (Send-AccuMail @splat) {
                foreach ($test in $client.TestArray) {
                    $prmAcc.Value = $test.ACCESSION
                    $prmDt.Value = "{0:yyyyMMdd HHmmss}" -f [datetime]::Now
                    $cmd.ExecuteNonQuery()
                }
            }
            Write-LogInfo ("{0} sent" -f $client.Name)
        }
    }

    [string] InterpretResult ([string] $Result) {
        $ResultDict = @{
            'ND'  = 'Not Detected' 
            'ON'  = 'Ongoing' 
            'INV' = 'Invalid' 
            'D'   = 'Detected' 
            'IN'  = 'Indeterminate' 
        }
        return $ResultDict.Keys -contains $Result ? $ResultDict.$Result : "Error"
    }

    [string] GetClientList () {
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
}

Class Test {
    [string] ${ACCESSION}
    [string] ${Final Report Date}
    [string] ${First Name}
    [string] ${Last Name}
    [string] ${Middle Name}
    [string] ${DOB}
    [string] ${Client ID}
    [string] ${Client Name}
    [string] ${Phys Name}
    [string] ${Test Code}
    [string] ${Test Name}
    [string] ${Test Result}
    [string] ${Patient Address}
    [string] ${Patient City}
    [string] ${Patient State}
    [string] ${Patient Zip}
    [string] ${Patient Phone}
}

class Client {
    [string]$Name
    [string[]]$To
    [string[]]$Cc
    [string[]]$Bcc
    [Test[]]$TestArray

    [string]$Header = @"
<style type="text/css">
table {
	border-collapse: collapse;
    font-family: Tahoma, Geneva, sans-serif;
}
table td {
	padding: 15px;
}
table thead td {
	background-color: #54585d;
	color: #ffffff;
	font-weight: bold;
	font-size: 13px;
	border: 1px solid #54585d;
}
table tbody td {
	color: #636363;
	border: 1px solid #dddfe1;
}
table tbody tr {
	background-color: #f9fafb;
}
table tbody tr:nth-child(odd) {
	background-color: #ffffff;
}
</style>
"@

    [string] GetHtmlBody () {
        $table = $this.TestArray | ConvertTo-Html -Fragment 
        return ConvertTo-Html -Body "$table" -Head $this.Header
    }
}

class Rep {
    [string]$Name
    [string[]]$To
    [string[]]$Cc
    [string[]]$Bcc
    [Test[]]$TestArray

    [string]$Header = @"
<style type="text/css">
table {
    border-collapse: collapse;
    font-family: Tahoma, Geneva, sans-serif;
}
table td {
    padding: 15px;
}
table thead td {
    background-color: #54585d;
    color: #ffffff;
    font-weight: bold;
    font-size: 13px;
    border: 1px solid #54585d;
}
table tbody td {
    color: #636363;
    border: 1px solid #dddfe1;
}
table tbody tr {
    background-color: #f9fafb;
}
table tbody tr:nth-child(odd) {
    background-color: #ffffff;
}
</style>
"@
    
    [string] GetHtmlBody () {
        $table = $this.TestArray | ConvertTo-Html -Fragment 
        return ConvertTo-Html -Body "$table" -Head $this.Header
    }
    
}
