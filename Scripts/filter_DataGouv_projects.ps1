# Retain CE projects only (-Match)
$filtering_column1 = "Projet.Code_Decision_ANR"
$filtering_expression1 = "^ANR-\d{2}-CE" 

# Retain only those with "artificial" and then "intelligence" the English abstract (-Match),
# not too far away, though
$filtering_column2 = "Projet.Resume.anglais"
$filtering_expression2 = "artificial.{1,20}intelligence" 

# Exclude those projects from CE23 (-NotMatch)
$filtering_column3 = "Projet.Code_Decision_ANR"
$filtering_expression3 = "-CE23-"


# You may wish to expand the list as willing! :-)  
# Make sure to add corresponding Where-Object commands below, with the right filtering option

#####################################################
# Phase 1: Download the DataGouv data into a local file

$url = "https://www.data.gouv.fr/fr/datasets/r/ecb8ec1b-a9e8-4ce0-8891-010ca1ca808f"

$data = "data.xlsx"
$path = "${HOME}/${data}"

$test = Test-Path -Path "$path" -PathType Leaf  # Leaf spécifie qu'on cherche un fichier, pas un répertoire

if ($test) {
    Write-Host("File ${path} already exists")
}
else {
    Write-Host("Loading URL ${url}")
    $WebClient = New-Object net.webclient
    $WebClient.DownloadFile($url, $path)
}

#####################################################
# Phase 2: Install the ImportExcel Module

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
# Phase 3: Import the local data file into a PowerShell object, and process it

Write-Host("Loading Excel file ${path}")

$excel = Import-Excel -Path "$path"
$result = $excel 

#####################################################
# Phase 4: Filter the result, using a sequence of positive (-Match) and negative (-NotMatch) patterns 
# https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/where-object

# Positive filtering (-Match)
$result = $result | Where-Object -Property "$filtering_column1" -Match "$filtering_expression1"

# Positive filtering (-Match)
$result = $result | Where-Object -Property "$filtering_column2" -Match "$filtering_expression2"

# Negative filtering (-NotMatch)
$result = $result | Where-Object -Property "$filtering_column3" -NotMatch "$filtering_expression3"

#####################################################
# Phase 5: Export the result as an Excel file and show it up

$result | Export-Excel -Show -AutoSize -AutoFilter -FreezeTopRow