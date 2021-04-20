$a0 = @{x = 1; y = 2 }
$a1 = @{x = 3; y = 4 }

$a = @($a0, $a1)

# $a | Get-Member
foreach ($line in $a) { Write-Host "==> $($line.keys)"; $line["z"] = $line.x + $line.y }
$a
# $a | ForEach-Object { $_.keys() }