$a0 = @{x = 1; y = 2 }
$a1 = @{x = 3; y = 4 }

$a = @($a0, $a1)

$a | foreach { $_.keys() }