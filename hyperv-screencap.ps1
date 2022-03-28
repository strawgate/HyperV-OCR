import-module .\hyperv-keyboard.psm1 -force

$Keyboard = Connect-VMKeyboard -ComputerName whitestone -VMName prd-rpt-srv-01

Send-VMKeyboardChar -KeyboardConnection $Keyboard -Key "T"