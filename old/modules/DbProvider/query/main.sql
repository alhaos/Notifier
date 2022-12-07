select
    l.Accession,
    l.[Final Report Date],
    RTRIM(l.[Patient First Name]) as "First Name",
    RTRIM(l.[Patient Last Name]) as "Last Name",
    l.MI as "Middle Name",
    l.DOB,
    l.[Client ID],
    l.[Client Name],
    l.[Phys  Name] as "Phys Name",
    l.[Test Code],
    l.[Test Name],
    RTRIM(LTRIM(l.Result)) as "Test Result",
    l.[Patient Address],
    l.[Patient City],
    l.[Patient State],
    l.[Patient Zip],
    l.[Patient Phone]
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