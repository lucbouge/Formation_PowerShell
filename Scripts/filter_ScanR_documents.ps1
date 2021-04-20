$query = "(Luc Bougé)|(L Bougé)"

$body = '{
    "page": 0,
    "pageSize": 1000,
    "query": "Luc Bougé",
    "searchFields": [
      "fullName"
    ],
    "sourceFields": [
      "id"
    ]
  }' 


$uri = "https://scanr-api.enseignementsup-recherche.gouv.fr/api/v2/persons/search"

# https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/invoke-webrequest

# https://davidhamann.de/2019/04/12/powershell-invoke-webrequest-by-example/

$r = Invoke-WebRequest -URI $uri -Body $body -Method 'POST' -ContentType 'application/json; charset=utf-8'

echo $r

# $result = ($r.Content | ConvertFrom-Json).response

# $numFound = $result.numFound
# Write-Host "numFound: $($result.numFound)"

# $result.docs | Export-Excel -Show -AutoSize -AutoFilter -FreezeTopRow
