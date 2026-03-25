
$bytes = [System.IO.File]::ReadAllBytes("C:\Users\Darin.Hochwender\_projects\AFP-Datastream-Assistant-ADA-V2.0\AFPEngineer_Explorer\AFP Files\DEV_IDS_150120261659_0092_Print.afp")
for($i=0; $i -lt $bytes.Length-8; $i++) {
    if($bytes[$i] -eq 0x5A) {
        $l = ($bytes[$i-2] -shl 8) -bor $bytes[$i-1]
        $type = "{0:X2}{1:X2}{2:X2}" -f $bytes[$i+1], $bytes[$i+2], $bytes[$i+3]
        if($type -eq "D3AB8A") {
            $hex = ""
            for($j=8; $j -lt [Math]::Min($l,200); $j++) { $hex += "{0:X2} " -f $bytes[$i+$j] }
            Write-Host "MCF len $l`: $hex"
        }
        if($type -eq "D3AFC3") {
            $hex = ""
            for($j=8; $j -lt [Math]::Min($l,200); $j++) { $hex += "{0:X2} " -f $bytes[$i+$j] }
            Write-Host "IOB len $l`: $hex"
        }
        $i += $l-3
    }
}

