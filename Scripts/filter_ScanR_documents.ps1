##########################################################
# Modifiy this query as wanted

$query = "(Luc Bouge)|(L Bouge)"

$sourceFields = @("id", "title", "authors")

##########################################################
# You might also modify the body of the request 

$body = @{
  page         = 0
  pageSize     = 1000
  query        = $query
  searchFields = @("authors.fullName")
  sourceFields = $sourceFields
}

##########################################################

$uri = "https://scanr-api.enseignementsup-recherche.gouv.fr/api/v2/publications/search"

# https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/invoke-webrequest
# https://davidhamann.de/2019/04/12/powershell-invoke-webrequest-by-example/

$body_json = $body | ConvertTo-Json

$r = Invoke-WebRequest -URI $uri -Body $body_json -Method 'POST' -ContentType 'application/json; charset=utf-8'

##########################################################

$results = ($r.Content | ConvertFrom-Json).results.value

$results = $results | Select-Object -Property $sourceFields

##########################################################
function get_author_data($author) {
  $fullName = $author.fullName
  $id = $author.person.id
  if (!$id) {
    $id = "none"
  }
  return "$fullName ($id)"
}
function replace_authors($line) {
  $line.authors = $line.authors | ForEach-Object { $author = $_; get_author_data($author) } 
}
function replace_title($line) {
  $line.title = $line.title.default
}
function linearize($line, $field) {
  if ($line.$field -is [array]) {
    $line.$field = $line.$field -join '; '
  }
}

$results | ForEach-Object { 
  $line = $_
  replace_authors($line)
  replace_title($line)
  foreach ($field in $sourceFields) { linearize $line $field; }
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
        
$excel_package = $results | Export-Excel -PassThru -Now -WorksheetName "Query" -FreezeTopRowFirstColumn -BoldTopRow 
$excel_package."Query".Cells.AutoFitColumns(5, 30) 
Export-Excel -ExcelPackage $excel_package -Show