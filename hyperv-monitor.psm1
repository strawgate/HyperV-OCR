
Class VMScreen {
    [string] $VMName
    [System.Drawing.Bitmap] $Bitmap

    VMScreen ([string] $VMName, [System.Drawing.Bitmap] $Bitmap) {
        $this.VMName = $VMName
        $this.Bitmap = $Bitmap
    }

    [System.Drawing.Bitmap] GetBitmap() { 
        return $this.Bitmap
    }

    SaveBitmap ([string] $Path) {
        $this.Bitmap.Save($Path)
    }
}

Function Get-VMScreenBitmap {
    [OutputType([VMScreen])]
    <#
    .SYNOPSIS
    Connects to a Hyper-V Server and takes a screenshot of the current video output of the desired virtual machine
    
    .DESCRIPTION
    Connects to a Hyper-V Server and takes a screenshot of the current video output of the desired virtual machine
    
    .PARAMETER Computername
    The name of the Hyper-V host with WinRM / Remote WMI enabled
    
    .PARAMETER VMName
    Parameter The name of the Hyper-V Guest to take a screenshot of
    
    .EXAMPLE
    Get-VMScreenBitmap -Computername MyHyperVHost -VMName MyVM
    
    .NOTES
    General notes
    #>
    param (
        [string] $Computername = "localhost",
        [parameter(Mandatory=$true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
        [string] $VMName
    )

    begin {
    }

    process {
        $MSVMVSMS = Get-CimInstance -Namespace root\virtualization\v2 -Class Msvm_VirtualSystemManagementService -ComputerName $Computername

        $MSVMCS   = Get-CimInstance -Namespace root\virtualization\v2 -Class Msvm_ComputerSystem                 -ComputerName $Computername -Filter "ElementName='$VMName'"
        $MSVMVH   = Get-CimAssociatedInstance -InputObject $MSVMCS -ResultClassName "Msvm_VideoHead"
        
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
    
        write-output ([VMScreen]::new($VMName, $BitMap))
    }

    end {}
}
<#
Function Save-VMScreenBitmap {
    param (
        [string] $Computername = "localhost",
        [string] $VMName,
        [string] $Outfile
    )

    begin {
        $MSVMVSMS = Get-CimInstance -Namespace root\virtualization\v2 -Class Msvm_VirtualSystemManagementService -ComputerName $Computername
    }

    process {
        $MSVMCS   = Get-CimInstance -Namespace root\virtualization\v2 -Class Msvm_ComputerSystem                 -ComputerName $Computername -Filter "ElementName='$VMName'"
        $MSVMVH   = Get-CimAssociatedInstance -InputObject $MSVMCS -ResultClassName "Msvm_VideoHead"
        
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
    
        $BitMap.Save($Outfile)
    }

    end {}
}
#>