$url = "https://www.data.gouv.fr/fr/datasets/r/ecb8ec1b-a9e8-4ce0-8891-010ca1ca808f"

#####################################################
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

Write-Host("Loading Excel file ${path}")

$excel = Import-Excel -Path "$path"

# https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/where-object

$result = $excel | Where-Object -Property "Projet.Code_Decision_ANR" -Match "^ANR-21" 

$result | Export-Excel -Show -AutoSize -AutoFilter -FreezeTopRow