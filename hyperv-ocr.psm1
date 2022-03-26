$Script:TesseractExe = ""

Function Set-TesseractExecutable {
    param (
        $Path
    )

    $Script:TesseractExe = $Path
}

Function Get-VMScreenText {
    param (
        $Computername = "localhost",
        $VMName
    )

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

    $Result = & $TesseractExe -l $Language $Path stdout

    return $Result
}

Function Save-VMScreenBitmap {
    param (
        $Computername = "localhost",
        $VMName,
        $Path
    )

    $MSVMVSMS = Get-WmiObject -Namespace root\virtualization\v2 -Class Msvm_VirtualSystemManagementService -ComputerName $Computername
    $MSVMCS   = Get-WmiObject -Namespace root\virtualization\v2 -Class Msvm_ComputerSystem                 -ComputerName $Computername -Filter "ElementName='$VMName'"

    $video = $MSVMCS.GetRelated("Msvm_VideoHead")

    $xResolution = $video.CurrentHorizontalResolution
    $yResolution = $video.CurrentVerticalResolution

    # Get the screenshot
    $image = $MSVMVSMS.GetVirtualSystemThumbnailImage($MSVMCS, $xResolution, $yResolution).ImageData

    # Transform into bitmap
    $Format = "Format16bppRgb565"
    $BitMap = New-Object System.Drawing.Bitmap    -Args $xResolution,$yResolution,$Format
    $Rect   = New-Object System.Drawing.Rectangle -Args 0,0,$xResolution,$yResolution

    $BmpData = $BitMap.LockBits($Rect,"ReadWrite",$Format)
    
    [System.Runtime.InteropServices.Marshal]::Copy($Image, 0, $BmpData.Scan0, $BmpData.Stride * $BmpData.Height)

    $BitMap.UnlockBits($BmpData)

    $BitMap.Save($Path)
}
