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

function f($authors) {
  return $authors.length
}


$results = ($r.Content | ConvertFrom-Json).results
$numbers = $results | ForEach-Object { f($_.value.authors) }
$results | Add-Member -MemberType NoteProperty -Name Numbers -Value $numbers
$results | Get-Member

$numbers = $results.numbers 
$id = $results.id
$title = $results.title

$results


# Export-Excel -Show -AutoSize -AutoFilter -FreezeTopRow
