##########################################################
# Modifiy this query as wanted

$query = "(Luc Bouge)|(L Bouge)"

##########################################################
# You might also modify the body of the request 

$body = @{
  page         = 0
  pageSize     = 1000
  query        = $query
  searchFields = @("authors.fullName")
  sourceFields = @("authors", "id", "title")
}

##########################################################

$uri = "https://scanr-api.enseignementsup-recherche.gouv.fr/api/v2/publications/search"

# https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/invoke-webrequest
# https://davidhamann.de/2019/04/12/powershell-invoke-webrequest-by-example/

$body_json = $body | ConvertTo-Json

$r = Invoke-WebRequest -URI $uri -Body $body_json -Method 'POST' -ContentType 'application/json; charset=utf-8'

##########################################################

$results = ($r.Content | ConvertFrom-Json).results.value
$results = $results | Select-Object -Property id, title, authors

##########################################################

$results | ForEach-Object { $_.title = $_.title.default }

function f($authors) {
  $fullNames = $authors | ForEach-Object { "$($_.fullName) ($($_.person.id))" }
  return $fullNames -join '; '
}

$results | ForEach-Object { $_.authors = f($_.authors) }

##########################################################

$results | Export-Excel -Show -AutoSize -AutoFilter -FreezeTopRow