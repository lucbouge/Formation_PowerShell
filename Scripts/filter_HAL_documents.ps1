##########################################################
# Modifiy this query as wanted

$query = "anrProjectReference_t:ANR-19-CE23-"

$fields = @("docid", "uri_s", "anrProjectReference_s", "authFullName_s", "publicationDateY_i", "title_s")

##########################################################
# You might also modify the body of the request 

$body = @{
  wt   = "json"
  rows = 1000
  q    = $query
  fl   = $fields -join ","
}

##########################################################

$uri = "https://api.archives-ouvertes.fr/search"

# https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/invoke-webrequest
# https://davidhamann.de/2019/04/12/powershell-invoke-webrequest-by-example/


$r = Invoke-WebRequest -URI $uri -Body $body -Method 'GET' 

##########################################################

$results = ($r.Content | ConvertFrom-Json).response.docs

$properties = @("docid", "uri_s", "anrProjectReference_s", "authFullName_s", "publicationDateY_i", "title_s")

$results = $results | Select-Object -Property $properties

###########################################################
function linearize($line, $field) {
  if ($line.$field -is [array]) {
    $line.$field = $line.$field -join '; '
  }
}

$results | ForEach-Object { 
  $line = $_
  foreach ($field in $fields) { linearize $line $field; }
}

##########################################################
# The -Now switch is a shortcut that automatically creates a temporary file, 
# enables "AutoSize", "TableName" and "Show", and opens the file immediately.
        
$excel_package = $results | Export-Excel -PassThru -Now -WorksheetName "Query"-FreezeTopRowFirstColumn -BoldTopRow 
$excel_package."Query".Cells.AutoFitColumns(20, 20) 
Export-Excel -ExcelPackage $excel_package -Show