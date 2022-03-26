$Script:TesseractExe = ""
$Script:TesseractLang = ""

$ErrorActionPreference = "stop"
$ErrorView = 'DetailedView'

Function Load-TesseractDefaults {
    $Executables = Get-ChildItem -Recurse -File -Filter "*.exe"

    $Tesseract = @($Executables).Where{$_.name -like "*tesseract-master.exe" -or $_.name -eq "tesseract.exe"}

    if ($Tesseract.Count -gt 0) {
        Set-TesseractExecutable -Path $Tesseract[0]
    } else {
        throw "Could not automatically locate Tesseract Executable. Use Set-TesseractExecutable set the executable path."
    }

    $Languages = @(Get-ChildItem -Recurse -File -Filter "*.traineddata")

    if ($Languages.Count -eq 1) {
        Set-TesseractLanguage -Language $Languages[0].BaseName
    } elseif ($Languages.Count -gt 1) {
        throw "Could not automatically locate Tesseract Language. Too many language files found. Use Set-TesseractLanguage to set the language."
    }
    else {
        throw "Could not automatically locate Tesseract Language. No language files found. Use Set-TesseractLanguage set the language."
    }
}

Function Get-TesseractExecutable { return $Script:TesseractExe}
Function Get-TesseractLanguage { return $Script:TesseractLanguage}

Function Set-TesseractExecutable {
    param (
        $Path
    )

    $Script:TesseractExe = $Path
}

Function Set-TesseractLanguage {
    param (
        $Language
    )

    $LanguageFileName = ($Language + ".traineddata")
    $LanguageFilePath = Join-Path $PSScriptRoot $LanguageFileName

    $Script:TesseractLanguage = $Language

    if (! (test-path $LanguageFilePath)) {
        throw "Could not locate language file $LanguageFilePath"
    }
}


Function Assert-TesseractExecutable {
    if (! (Test-TesseractExecutable)) {
        Throw "Could not find Tesseract Executable. Set location of Tesseract Executable using Set-TesseractExecutable."
    }
}

Function Assert-TesseractLanguage {
    if (! (Test-TesseractLanguage)) {
        Throw "Tesseract Language is not set. Set the Tesseract Language using Set-TesseractLanguage."
    }
}


Function Test-TesseractLanguage {
    if ([string]::IsNullOrWhiteSpace($Script:TesseractLang)) {
        return $False
    }

    return $True
}

Function Test-TesseractExecutable {
    if (! (Test-Path $Script:TesseractExe)) {
        return $False
    }

    return $True
}

Function Get-VMScreenText {
    param (
        $Computername = "localhost",
        $VMName
    )
    
    Assert-TesseractExecutable

    $TempPath = [System.IO.Path]::GetTempFileName()

    Save-VMScreenBitmap -Computername $Computername -VMName $VMName -Path $TempPath

    $Text = ConvertImage-ToText -Path $TempPath

    Remove-Item $TempPath

    return $Text
}

Function ConvertImage-ToText {
    param (
        [string] $Path,
        [string] $Language = "eng"
    )
    Assert-TesseractExecutable

    $Result = & $TesseractExe -l $Language $Path stdout

    return $Result
}

Function Save-VMScreenBitmap {
    param (
        $Computername = "localhost",
        $VMName,
        $Path
    )

    $MSVMVSMS = Get-CimInstance -Namespace root\virtualization\v2 -Class Msvm_VirtualSystemManagementService -ComputerName $Computername
    $MSVMCS   = Get-CimInstance -Namespace root\virtualization\v2 -Class Msvm_ComputerSystem                 -ComputerName $Computername -Filter "ElementName='$VMName'"
    $MSVMVH   = Get-CimAssociatedInstance -InputObject $MSVMCS -ResultClassName "Msvm_VideoHead"

    $xResolution = $MSVMVH.CurrentHorizontalResolution
    $yResolution = $MSVMVH.CurrentVerticalResolution

    $image = Invoke-CIMMethod -InputObject $MSVMVSMS -MethodName "GetVirtualSystemThumbnailImage" -Arguments @{
        "TargetSystem"  = $MSVMCS
        "WidthPixels"   = $xResolution
        "HeightPixels"  = $yResolution
    }

    $image = $image.ImageData

    # Transform into bitmap
    $Format = "Format16bppRgb565"
    $BitMap = New-Object System.Drawing.Bitmap    -Args $xResolution,$yResolution,$Format
    $Rect   = New-Object System.Drawing.Rectangle -Args 0,0,$xResolution,$yResolution

    $BmpData = $BitMap.LockBits($Rect,"ReadWrite",$Format)
    
    [System.Runtime.InteropServices.Marshal]::Copy($Image, 0, $BmpData.Scan0, $BmpData.Stride * $BmpData.Height)

    $BitMap.UnlockBits($BmpData)

    $BitMap.Save($Path)
}
