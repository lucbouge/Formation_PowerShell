$query = "(Luc Bouge)|(L Bouge)"

$body = @{
  page         = 0
  pageSize     = 1000
  query        = $query
  searchFields = @("authors.fullName")
  sourceFields = @("authors", "id", "title.default")
}

$uri = "https://scanr-api.enseignementsup-recherche.gouv.fr/api/v2/publications/search"

# https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/invoke-webrequest

# https://davidhamann.de/2019/04/12/powershell-invoke-webrequest-by-example/

$body_json = $body | ConvertTo-Json

$r = Invoke-WebRequest -URI $uri -Body $body_json -Method 'POST' -ContentType 'application/json; charset=utf-8'

$results = ($r.Content | ConvertFrom-Json).results

$lines = foreach ($line in $results.value) { 
  @{id      = $line.id
    authors = $line.authors
    title   = $line.title.default 
  }
}

$lines

# $lines | Export-Excel -Show -AutoSize -AutoFilter -FreezeTopRow
