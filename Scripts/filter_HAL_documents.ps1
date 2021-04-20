##########################################################
# Modifiy this query as wanted

$query = "anrProjectReference_t:ANR-19-CE23-"

##########################################################
# You might also modify the body of the request 

$body = @{
  wt   = "json"
  rows = 1000
  q    = $query
  fl   = "docid, uri_s, anrProjectReference_s, authFullName_s, publicationDateY_i, title_s"
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

# ##########################################################
function f($strings) {
  return $strings -join '; '
}

$results | ForEach-Object { 
  $_.authFullName_s = f($_.authFullName_s); 
  $_.anrProjectReference_s = f($_.anrProjectReference_s)
  $_.title_s = f($_.title_s)
}
# ##########################################################

$results | Export-Excel -Show -AutoSize -AutoFilter -FreezeTopRow