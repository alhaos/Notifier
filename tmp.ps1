$data = @"
{
    "Name": "Irina",
    "To": null,
    "Cc": [
      "irina.kanevsky@accureference.com",
      "helen.said-akl@accureference.com"
    ],
    "Bcc": [
      "interfaces@accureference.com"
    ],
    "TestArray": [
      {
        "ACCESSION": "2233135115",
        "Final Report Date": "2022-11-28 00:00:00",
        "First Name": "RENEE",
        "Last Name": "VINS",
        "Middle Name": " ",
        "DOB": "07/27/1947",
        "Client ID": "38203",
        "Client Name": "REGENCY GARDENS PA       ",
        "Phys Name": "ERHAN KUCUK              ",
        "Test Code": "950Z",
        "Test Name": "SARS CoV-2, SWAB (PCR)",
        "Test Result": "Detected",
        "Patient Address": "2 SPRINT ST",
        "Patient City": "HIGHLAND LAKES",
        "Patient State": "NJ",
        "Patient Zip": "07422",
        "Patient Phone": "(   )   -    "
      }
    ]
  }
"@ | ConvertFrom-Json


$data.TestArray | ConvertTo-Html -Fragment