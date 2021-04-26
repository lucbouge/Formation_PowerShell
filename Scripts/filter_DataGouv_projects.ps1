function test($line) {

    # Retain CE projects only (-match)
    #
    $test1 = $line."Projet.Code_Decision_ANR" -match "^ANR-20" 

    # Retain only those with "artificial" and then "intelligence" the English abstract (-match),
    # not too far away, though
    #
    $test2 = $line."Projet.Resume.anglais" -match "artificial.{1,20}intelligence" 

    # Exclude those projects from CE23 (-notmatch)
    #
    $test3 = $line."Projet.Code_Decision_ANR" -notmatch "-CE23-" 

    # Build the logical outcome
    #
    $test = $test1 -and $test2 -and $test3
    
    # Return the value
    return $test
}

# You may wish to expand the list as willing! :-)  

#####################################################
# Phase 1: Download the DataGouv data into a local file

$url = "https://www.data.gouv.fr/fr/datasets/r/b0ce2c54-68a9-4493-83d7-0721a391fb4f"

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

#####################################################
# Phase 4: Filter the Excel file with the test function
# https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/where-object

$results = $excel | Where-Object { test($_) } 

#####################################################
# Phase 5: Export the result as an Excel file and show it up
# The -Now switch is a shortcut that automatically creates a temporary file, 
# enables "AutoSize", "TableName" and "Show", and opens the file immediately.
        
$excel_package = $results | Export-Excel -PassThru -Now -WorksheetName "Query"-FreezeTopRowFirstColumn -BoldTopRow 
$excel_package."Query".Cells.AutoFitColumns(5, 30) 
Export-Excel -ExcelPackage $excel_package -Show