param (
    $ComputerName,
    $VMName
)

#Import-Module (Join-Path $PSScriptRoot ".\keyboard.psm1") -Force
Import-Module (Join-Path $PSScriptRoot ".\hyperv-ocr.psm1") -Force

Load-TesseractDefaults

Save-VMScreenBitmap -ComputerName $ComputerName -VMName $VMName -Path screen.bmp

Get-VMScreenText -ComputerName $ComputerName -VMName $VMName