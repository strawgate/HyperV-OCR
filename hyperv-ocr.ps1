param (
    $ComputerName,
    $VMName
)

Import-Module (Join-Path $PSScriptRoot ".\hyperv-ocr.psm1") -Force

Load-TesseractDefaults

Get-VMScreenText -ComputerName $ComputerName -VMName $VMName