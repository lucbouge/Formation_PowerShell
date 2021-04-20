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
  $line.$field = $line.$field -join '; '
}

$list_fields = ("authFullName_s", "anrProjectReference_s", "title_s")

$results | ForEach-Object { 
  $line = $_
  foreach ($field in $list_fields) { linearize $line $field; }
}

###########################################################

$results | Export-Excel -Show -AutoSize -AutoFilter -FreezeTopRow