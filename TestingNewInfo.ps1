Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue 
$connection= new-object system.data.sqlclient.sqlconnection
$Connection.ConnectionString ="server=LMSAG;database=LMSFILLIB;trusted_connection=true"
Write-host "connection information:"
$connection
$connection.open()
$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
$SqlQuery = Get-Content "C:\Users\pgirdhar\GetQueryNew.txt" <# SQL Query to get data from RIPPE #>
$SqlCmd.CommandText = $SqlQuery
$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
$SqlAdapter.SelectCommand = $SqlCmd
$SqlCmd.CommandTimeout = 0
$SqlCmd.Connection = $connection 
$DataSet = New-Object System.Data.DataSet 
$SqlAdapter.Fill($DataSet)
$SqlAdapter.Dispose()
$connection.Close() 
$data = $DataSet.Tables[0]
<#$folderLocation= "C:\Users\pgirdhar\MicrosoftFlow\Beazley" # - Folder Location #>
$folderLocation= "\\ODRIVEFS\Administrative Departments\Information Services\Pratik\ClientFolderTest‎" # - Folder Location
foreach ($Row in $data)
{
$ClientName = $Row["CMNAME"].trim()
$ClientCode = $Row["CLNT"].trim()
$matterCode = $Row["mm"].trim()
$matterdescription = $Row["mmdesc"].trim()
$mm023 = $Row["MMC023"].trim()
$ClientMatterCode = $ClientName+"-"+$ClientCode
<#$matterCodeFull = $matterdescription + "-" + $matterCode #>
$matterCodeFull = $matterCode + "-" + $matterdescription
$matterCodeFull = $matterCodeFull -replace '&','' 
$matterCodeFull = $matterCodeFull -replace '/',''
$matterCodeFull = $matterCodeFull -replace ',',''
$matterCodeFull = $matterCodeFull -replace '"',''
$matterCodeFull = $matterCodeFull -replace ':',''
$matterCodeFull = $matterCodeFull.trim()
$NewFolderLocation = $folderLocation + "\$ClientMatterCode"
if (-not (Test-Path -LiteralPath $NewFolderLocation)){
   try {
      New-Item -Path $NewFolderLocation -ItemType Directory -ErrorAction Stop | Out-Null #-Force
     
      $NewMatterLocationClient = $NewFolderLocation + "\$matterCodeFull"
      New-Item  -Path $NewMatterLocationClient -ItemType Directory -ErrorAction Stop | Out-Null #-Force
      <#Write-Host "Create New Client Folder $NewFolderLocation" #>
      <#Write-Host "Create New Matter Folder $NewMatterLocationClient inside $NewFolderLocation" #>
      
      if ($mm023 -eq '07000') {

      $FilingsDocket = $NewMatterLocationClient + "\01 Filings-Docket"
      New-Item -Path $FilingsDocket -ItemType Directory -ErrorAction Stop | Out-Null 
      $Discovery = $NewMatterLocationClient + "\02 Discovery"
      New-Item -Path $Discovery -ItemType Directory -ErrorAction Stop | Out-Null
      $Disposition = $NewMatterLocationClient + "\03 Depositions"
      New-Item -Path $Disposition -ItemType Directory -ErrorAction Stop | Out-Null
      $Correspondence = $NewMatterLocationClient + "\04 Correspondence"
      New-Item -Path $Correspondence -ItemType Directory -ErrorAction Stop | Out-Null
      $Experts = $NewMatterLocationClient + "\05 Experts"
      New-Item -Path $Experts -ItemType Directory -ErrorAction Stop | Out-Null
      $Working = $NewMatterLocationClient + "\06 Working"
      New-Item -Path $Working -ItemType Directory -ErrorAction Stop | Out-Null
      $Trial = $NewMatterLocationClient + "\07 Trial"
      New-Item -Path $Trial -ItemType Directory -ErrorAction Stop | Out-Null
      $Background = $NewMatterLocationClient + "\08 Background"
      New-Item -Path $Background -ItemType Directory -ErrorAction Stop | Out-Null
      $Admin = $NewMatterLocationClient + "\09 Admin"
      New-Item -Path $Admin -ItemType Directory -ErrorAction Stop | Out-Null
      $PST = $NewMatterLocationClient + "\ZZZ_PST"
      New-Item -Path $PST -ItemType Directory -ErrorAction Stop | Out-Null
      
    }

      if ($mm023 -eq '22000') {

      $Subpoenas = $NewMatterLocationClient + "\01 Subpoenas and Correspondence"
      New-Item -Path $Subpoenas -ItemType Directory -ErrorAction Stop | Out-Null
      $ClientDocument = $NewMatterLocationClient + "\02 Client Documents"
      New-Item -Path $ClientDocument -ItemType Directory -ErrorAction Stop | Out-Null
      $Timelines = $NewMatterLocationClient + "\03 Timelines"
      New-Item -Path $Timelines -ItemType Directory -ErrorAction Stop | Out-Null
      $LegalResearch = $NewMatterLocationClient + "\04 Legal Research"
      New-Item -Path $LegalResearch -ItemType Directory -ErrorAction Stop | Out-Null
      $LitigationFiling = $NewMatterLocationClient + "\05 Litigation Filings"
      New-Item -Path $LitigationFiling -ItemType Directory -ErrorAction Stop | Out-Null
      $LitigationDiscovery = $NewMatterLocationClient + "\06 Litigation Discovery"
      New-Item -Path $LitigationDiscovery -ItemType Directory -ErrorAction Stop | Out-Null
      $LitigationCorrespondence = $NewMatterLocationClient + "\07 Litigation Correspondence"
      New-Item -Path $LitigationCorrespondence -ItemType Directory -ErrorAction Stop | Out-Null
      $Experts = $NewMatterLocationClient + "\08 Experts"
      New-Item -Path $Experts -ItemType Directory -ErrorAction Stop | Out-Null
      $DocumentProductions = $NewMatterLocationClient + "\09 Document Productions"
      New-Item -Path $DocumentProductions -ItemType Directory -ErrorAction Stop | Out-Null
      $WorkingFiles = $NewMatterLocationClient + "\10 Working Files"
      New-Item -Path $WorkingFiles -ItemType Directory -ErrorAction Stop | Out-Null
      $11Trial = $NewMatterLocationClient + "\11 Trial"
      New-Item -Path $11Trial -ItemType Directory -ErrorAction Stop | Out-Null
      $ZZZ = $NewMatterLocationClient + "\ZZZ_PST"
      New-Item -Path $ZZZ -ItemType Directory -ErrorAction Stop | Out-Null
      
}
         
}

catch {
   Write-Error -Message "Unable to create directory '$NewFolderLocation'. Error was: $_" -ErrorAction Stop    
 }

}
else {
 "Directory Already existed" <# Message if Client Already present #>
  $NewMatterLocation = $NewFolderLocation + "\$matterCodeFull"
  if(-not (Test-Path -LiteralPath $NewMatterLocation)){

      try{
        New-Item -Path $NewMatterLocation -ItemType Directory -ErrorAction Stop | Out-Null #-Force
        Write-Host "Created new Matter Folder $NewMatterLocation"
        if ($mm023 -eq '07000') {
      $FilingsDocket = $NewMatterLocation + "\01 Filings-Docket"
      New-Item -Path $FilingsDocket -ItemType Directory -ErrorAction Stop | Out-Null
      $Discovery = $NewMatterLocation + "\02 Discovery"
      New-Item -Path $Discovery -ItemType Directory -ErrorAction Stop | Out-Null
      $Disposition = $NewMatterLocation + "\03 Depositions"
      New-Item -Path $Disposition -ItemType Directory -ErrorAction Stop | Out-Null
      $Correspondence = $NewMatterLocation + "\04 Correspondence"
      $Experts = $NewMatterLocation + "\05 Experts"
      New-Item -Path $Experts -ItemType Directory -ErrorAction Stop | Out-Null
      $Working = $NewMatterLocation + "\06 Working"
      New-Item -Path $Working -ItemType Directory -ErrorAction Stop | Out-Null
      $Trial = $NewMatterLocation + "\07 Trial"
      New-Item -Path $Trial -ItemType Directory -ErrorAction Stop | Out-Null
      $Background = $NewMatterLocation + "\08 Background"
      New-Item -Path $Background -ItemType Directory -ErrorAction Stop | Out-Null
      $Admin = $NewMatterLocation + "\09 Admin"
      New-Item -Path $Admin -ItemType Directory -ErrorAction Stop | Out-Null
      $PST = $NewMatterLocation + "\ZZZ_PST"
      New-Item -Path $PST -ItemType Directory -ErrorAction Stop | Out-Null
        }
        if ($mm023 -eq '22000'){
      $Subpoenas = $NewMatterLocation + "\01 Subpoenas and Correspondence"
      New-Item -Path $Subpoenas -ItemType Directory -ErrorAction Stop | Out-Null
      $ClientDocument = $NewMatterLocation + "\02 Client Documents"
      New-Item -Path $ClientDocument -ItemType Directory -ErrorAction Stop | Out-Null
      $Timelines = $NewMatterLocation + "\03 Timelines"
      New-Item -Path $Timelines -ItemType Directory -ErrorAction Stop | Out-Null
      $LegalResearch = $NewMatterLocation + "\04 Legal Research"
      New-Item -Path $LegalResearch -ItemType Directory -ErrorAction Stop | Out-Null
      $LitigationFiling = $NewMatterLocation + "\05 Litigation Filings"
      New-Item -Path $LitigationFiling -ItemType Directory -ErrorAction Stop | Out-Null
      $LitigationDiscovery = $NewMatterLocation + "\06 Litigation Discovery"
      New-Item -Path $LitigationDiscovery -ItemType Directory -ErrorAction Stop | Out-Null
      $LitigationCorrespondence = $NewMatterLocation + "\07 Litigation Correspondence"
      New-Item -Path $LitigationCorrespondence -ItemType Directory -ErrorAction Stop | Out-Null
      $Experts = $NewMatterLocation + "\08 Experts"
      New-Item -Path $Experts -ItemType Directory -ErrorAction Stop | Out-Null
      $DocumentProductions = $NewMatterLocation + "\09 Document Productions"
      New-Item -Path $DocumentProductions -ItemType Directory -ErrorAction Stop | Out-Null
      $WorkingFiles = $NewMatterLocation + "\10 Working Files"
      New-Item -Path $WorkingFiles -ItemType Directory -ErrorAction Stop | Out-Null
      $11Trial = $NewMatterLocation + "\11 Trial"
      New-Item -Path $11Trial -ItemType Directory -ErrorAction Stop | Out-Null
      $ZZZ = $NewMatterLocation + "\ZZZ_PST"
      New-Item -Path $ZZZ -ItemType Directory -ErrorAction Stop | Out-Null

        }
      }
      catch {
        Write-Error -Message "Unable to create Matter '$NewMatterLocation'. Error was: $_" -ErrorAction Stop

      }
  }
  else {
      "Matter Folder Already Existed" <# Message if Matter existed #>
  }
}
}

