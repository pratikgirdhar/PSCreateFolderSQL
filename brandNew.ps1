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
$folderLocation= "C:\Users\pgirdhar\MicrosoftFlow\Beazley" # - Folder Location 
foreach ($Row in $data)
{
$ClientName = $Row["CMNAME"].trim()
$ClientCode = $Row["CLNT"].trim()
$matterCode = $Row["mm"].trim()
$matterdescription = $Row["mmdesc"].trim()
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
      Write-Host "Create New Matter Folder $NewMatterLocationClient inside $NewFolderLocation"
      
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

