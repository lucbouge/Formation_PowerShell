function test($line) {
    # Param($line)
    $test1 = $line."Projet.Code_Decision_ANR" -match "^ANR-20" 
    $test2 = $line."Projet.Resume.anglais" -match "artificial.{1,20}intelligence" 
    $test3 = $line."Projet.Code_Decision_ANR" -notmatch "-CE23-" 
    return ($test1 -and $test2 -and $test3)
}

$data = "data.xlsx"
$path = "${HOME}/${data}"

$excel = Import-Excel -Path "$path"

#####################################################
   
# https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/where-object

$result = $excel | Where-Object { test($_) } 


#####################################################

$result | Export-Excel -Show -AutoSize -AutoFilter -FreezeTopRow