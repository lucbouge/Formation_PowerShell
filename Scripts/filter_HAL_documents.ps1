##########################################################
# Modifiy this query as wanted

$query = "anrProjectReference_t:""ANR-19-CE23-"""

$fields = @("docid", "uri_s", "anrProjectReference_s", "authFullName_s", "publicationDateY_i", "title_s")

##########################################################
# You might also modify the body of the request 

$body = @{
  wt   = "json"
  rows = 1000
  q    = $query
  fl   = $fields -join ", "
}

##########################################################

$uri = "https://api.archives-ouvertes.fr/search"

# https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/invoke-webrequest
# https://davidhamann.de/2019/04/12/powershell-invoke-webrequest-by-example/


$r = Invoke-WebRequest -URI $uri -Body $body -Method 'GET' 

##########################################################

$results = ($r.Content | ConvertFrom-Json).response.docs

$properties = @("docid", "uri_s", "anrProjectReference_s", "authFullName_s", "publicationDateY_i", "title_s")

$results = $results | Select-Object -Property $properties

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

#####################################################
# Install the ImportExcel Module
# https://github.com/dfinke/ImportExcel

$module = "ImportExcel"

$test = Get-Module -ListAvailable -Name "$module"

if ($test) {
  Write-Host("Module ${module} already installed")
}
else {
  Write-Host("Installing Module ${module}")
  Install-Module -Name "$module" -Scope CurrentUser
}

#####################################################
# Export the result as an Excel file and show it up
# The -Now switch is a shortcut that automatically creates a temporary file, 
# enables "AutoSize", "TableName" and "Show", and opens the file immediately.
        
$excel_package = $results | Export-Excel -PassThru -Now -WorksheetName "Query"-FreezeTopRowFirstColumn -BoldTopRow 
$excel_package."Query".Cells.AutoFitColumns(5, 30) 
Export-Excel -ExcelPackage $excel_package -Show