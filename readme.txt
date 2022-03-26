Requires Windows Powershell

Requires Tesseract
I grab latest artifact from here: https://ci.appveyor.com/project/zdenop/tesseract/history
Place it in a folder next to the library

Requires Tesseract Language Data: https://github.com/tesseract-ocr/tessdata
Place it next to hyperv-ocr.psm1

Usage:

```
Import-Module .\hyperv-ocr.psm1 -Force

Load-TesseractDefaults

Get-VMScreenText -ComputerName "hyperv-host" -VMName "vmname"
```