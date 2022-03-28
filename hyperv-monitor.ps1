import-module .\hyperv-monitor.psm1 -force

$Results = @("server1","server2") | get-vmscreenbitmap -Computername whitestone

$Results | % {$_.saveBitmap(".\" + $_.VMName + ".bmp")}