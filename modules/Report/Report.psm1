using namespace System.Data
using namespace System.Data.SqlClient
using namespace System.Data.SQLite

class Report {
    [SQLiteConnection] $LocalConnection
    [Rep[]]$Reps
    [Client[]]$Clients

    Report () {
        $this.LocalConnection = [SQLiteConnection]::new()
        $this.LocalConnection.ConnectionString = "Data source=./modules/DbProvider/database.db;Version=3"
        $this.LocalConnection.Open()
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

            if ($rep.TestArray.Count){
                $this.Reps += , $rep
            }
        }
        $dr.Close()
    }
 
    FillClients () {
        $cmdClient = $this.LocalConnection.CreateCommand()
        $cmdClient.CommandText = 'select distinct "CLIENT ID" from RAW_DATA'

        $cmdRepEmail = $this.LocalConnection.CreateCommand()
        $cmdRepEmail.CommandText = @"
 select et.NAME,
        ce.ADDRES
   from CLIENT_EMAIL ce
   join EMAIL_TYPE et
     on ce.TYPE_ID = et.ID
  where CLIENT_ID = @ID
"@
        $idParameter = [SQLiteParameter]::new("@ID", [DbType]::String)
        $cmdRepEmail.Parameters.Add($idParameter)

        $cmdTests = $this.LocalConnection.CreateCommand()
        $cmdTests.CommandText = @"
select rd.* from RAW_DATA rd
  join CLIENT c on rd."CLIENT ID" = c.ID
  join RESULT rt on rt.NAME = rd."TEST RESULT"
  join CLIENT_RESULT cr on c.id = cr.CLIENT_ID and rt.ID = cr.RESULT_ID
 where rd."CLIENT ID" = 38203
"@
        $cmdTests.Parameters.Add($idParameter)

        $dr = $cmdClient.ExecuteReader()
        while ($dr.Read()) {
            $Client = [Client]@{
                Name = $dr['CLIENT ID']
            }
            $idParameter.Value = $dr['ID']
            $drEmail = $cmdRepEmail.ExecuteReader()
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
}

class Rep {
    [string]$Name
    [string[]]$To
    [string[]]$Cc
    [string[]]$Bcc
    [Test[]]$TestArray
}

class Client {
    [string]$Name
    [string[]]$To
    [string[]]$Cc
    [string[]]$Bcc
    [Test[]]$TestArray

    [string]$Header = @"
    <title>Client report</title>
    <style>
        td,
        th {
            border-style: solid;
            border-color: black;
            border-width: 3px;
        }
    </style>
"@

    [string] GetHtmlBody () {
        $table = $this.TestArray | ConvertTo-Html -Fragment 
        return ConvertTo-Html -Body "$table" -Head $this.Header
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