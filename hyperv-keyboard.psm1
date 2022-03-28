
if ( $null -eq  ( 'Win32Functions.KeyboardScan' -as [type]) ) {
    Add-Type -Name KeyboardScan -Namespace Win32Functions -MemberDefinition @'

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern short VkKeyScanEx(char ch, IntPtr dwhkl);
'@
}

Function Validate-MSVMKeyboardCIMInstance {
	param (
		[CimInstance] $InputObject 
	)

	if ($InputObject -eq $Null) { return $false }
	if ($InputObject.CreationClassName -ne "Msvm_Keyboard") { return $False}

	return $true
}

Function Assert-MSVMKeyboardCIMInstance {
	param (
		[CimInstance] $InputObject 
	)

	if (-not (Validate-MSVMKeyboardCIMInstance -InputObject $InputObject)) {
		throw "Invalid CIM Instance. Requires a msvm_keyboard class."
	}
}

Function Connect-VMKeyboard {
	# Get and return a WMI object that represents a keyboard for the Hyper-V VM
	Param (
		[string] $ComputerName = "localhost",
		[string] $VMName
	)

	$MSVMCS = Get-CIMInstance -Query "select * from Msvm_ComputerSystem where ElementName = '$VMName'" -Namespace "root\virtualization\v2" -ComputerName $ComputerName
	[CimInstance] $MSVMKB = Get-CimAssociatedInstance -InputObject $MSVMCS -ResultClassName "Msvm_Keyboard"
	
	Assert-MSVMKeyboardCIMInstance -InputObject $MSVMKB

	return $MSVMKB
}


Function Hold-VMKeyboardShift {
	Param (
		[CimInstance] $KeyboardConnection
	)

	Assert-MSVMKeyboardCIMInstance -InputObject $KeyboardConnection

	# 160 is Right Shift
	$ResultObj = Invoke-CimMethod -InputObject $KeyboardConnection -MethodName "PressKey" -Arguments @{"keyCode" = 160}
	$result = $ResultObj.ReturnValue

	if ($Result -ne 0) { throw "Key entered returned a failure" }
}

Function Release-VMKeyboardShift {
	Param (
		[CimInstance] $KeyboardConnection
	)

	Assert-MSVMKeyboardCIMInstance -InputObject $KeyboardConnection

	# 160 is Right Shift
	$ResultObj = Invoke-CimMethod -InputObject $KeyboardConnection -MethodName "ReleaseKey" -Arguments @{"keyCode" = 160}
	$result = $ResultObj.ReturnValue
	
	if ($Result -ne 0) { throw "Key entered returned a failure" }
}

Function Get-KeyScan {
    param (
        [Char] $Key
    )

    $Result = [Win32Functions.KeyboardScan]::VkKeyScanEx("$Key", 0)

    $Flags = @{
        "Shift" = [bool] ($Result -band 256 -eq 256) # Shift is pressed
        "Code" = $Result
    }
}

Function Send-VMKeyboardChar {
	Param (
		[CimInstance] $KeyboardConnection,
		[Char] $Key
	)

	Assert-MSVMKeyboardCIMInstance -InputObject $KeyboardConnection

	# VirtualCodes are how systems know which keys were pressed. It's differnet from ascii and requires translation.
	# An important distinction is that in ascii A is a unique character but in virtual codes A is a combination of Shift + A.
	# This means we need to literally "hold shift" and then press our key and then "release shift"
	$VirtualKey = [Win32Functions.KeyboardScan]::VkKeyScanEx("$Key", 0)

	# This bit flag is set if we need to shift. Hold shift and then press the key.
	if ($VirtualKey -gt 256) {
        Write-Debug "Holding Shift before typing $Key"
		Hold-VMKeyboardshift -keyboardConnection $KeyboardConnection
	}

    Write-Debug "Typing $Key"

	$ResultObj = Invoke-CimMethod -InputObject $KeyboardConnection -MethodName "TypeKey" -Arguments @{"keyCode" = $VirtualKey}
	$ReturnValue = $ResultObj.ReturnValue

	if ($VirtualKey -gt 256) {
        Write-Debug "Releasing Shift after typing $Key"
		Release-VMKeyboardshift -keyboardConnection $KeyboardConnection
	}

	if ($ReturnValue -ne 0) { throw "Key entered returned a failure" }

}

Function Send-VMKeyboardString {
	# Linux VMs cannot receive characters from SendText so we have to use PressKey/SendKey which means one call per key
	Param (
		[CimInstance] $KeyboardConnection,
        
        # String to type
		[string] $String,

        # Delay between each character in milliseconds
		[int] $Delay = 0
	)

	Assert-MSVMKeyboardCIMInstance -InputObject $KeyboardConnection

	$String.ToCharArray().Foreach{
		write-debug "$_" -NoNewLine

		if ($Delay -ne 0) { Start-Sleep -milliseconds $Delay }

		Send-VMKeyboardChar -KeyboardConnection $KeyboardConnection -Key $_
	}
}

Function Send-VMCommand {
	<#
	.SYNOPSIS
	Types a text command to the virtual machine and hits enter. Optionally adds a delay between keystrokes and a duration to sleep after hitting enter.
	
	.DESCRIPTION
	Types a text command to the virtual machine and hits enter. Optionally adds a delay between keystrokes and a duration to sleep after hitting enter.
	
	.PARAMETER KeyboardConnection
	A connection to a Hyper-V VM Keyboard
	
	.PARAMETER Command
	The command to run on the Virtual Machine
	
	.PARAMETER Delay
	Duration to wait between keystrokes
	
	.PARAMETER Sleep
	Duration to wait after running the command and before returning to the caller.
	
	.EXAMPLE
	The following example will use systemctl to enable the Hyper-V service and wait one second before returning to the caller. 
	Send-VMCommand -KeyboardConnection $KeyboardConnection -Command "sudo systemctl enable hv_kvp_daemon" -Sleep 1000
	#>
	param (
		[CimInstance] $KeyboardConnection,
		[string] $Command,
		[int] $Delay = 10,
		[int] $Sleep = 0
	)

	Send-VMKeyboardString -KeyboardConnection $KeyboardConnection -String $Command -Delay $Delay
	Send-VMKeyboardChar -KeyboardConnection $KeyboardConnection -Key "`n"

    if ($Sleep -ne 0) { Start-Sleep -milliseconds $Sleep }
}


Export-ModuleMember -Function @("Send-VMCommand","Send-VMKeyboardString","Send-VMKeyboardChar")