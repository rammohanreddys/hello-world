<#
.SYNOPSIS
	This script identifies if any MAK license keys are installed and converts them to a KMS activated installation.
.DESCRIPTION
	The script is designed to be included in a Configuration Manager Task Sequence to remove any MAK license keys that are included within the Global image and replace them with the Global KMS license key, subsequently activating the product via the PwC Global KMS server.
	The script dot-sources the AppDeployToolkitMain.ps1 script which contains the logic and functions required to check applications, files and registry settings.
.PARAMETER DeployMode
	Specifies whether the installation should be run in Interactive, Silent, or NonInteractive mode. Default is: Interactive. Options: Interactive = Shows dialogs, Silent = No dialogs, NonInteractive = Very silent, i.e. no blocking apps. NonInteractive mode is automatically set if it is detected that the process is not user interactive.
.PARAMETER AllowRebootPassThru
	Allows the 3010 return code (requires restart) to be passed back to the parent process (e.g. SCCM) if detected from an installation. If 3010 is passed back to SCCM, a reboot prompt will be triggered.
.PARAMETER DisableLogging
	Disables logging to file for the script. Default is: $false.
.EXAMPLE
	Deploy-Application.ps1
.EXAMPLE
	Deploy-Application.ps1 -DeployMode 'Silent'
.EXAMPLE
	Deploy-Application.ps1 -AllowRebootPassThru -AllowDefer
.NOTES
	Toolkit Exit Code Ranges:
	60000 - 68999: Reserved for built-in exit codes in Deploy-Application.ps1, Deploy-Application.exe, and AppDeployToolkitMain.ps1
	69000 - 69999: Recommended for user customized exit codes in Deploy-Application.ps1
	70000 - 79999: Recommended for user customized exit codes in AppDeployToolkitExtensions.ps1
#>
[CmdletBinding()]
param (
	[Parameter(Mandatory = $false)]
	[ValidateSet('Interactive', 'Silent', 'NonInteractive')]
	[string]$DeployMode = 'Interactive',
	[Parameter(Mandatory = $false)]
	[switch]$AllowRebootPassThru = $true,
	[Parameter(Mandatory = $false)]
	[switch]$DisableLogging = $false
)

try
{
	## Set the script execution policy for this process
	try { Set-ExecutionPolicy -ExecutionPolicy 'ByPass' -Scope 'Process' -Force -ErrorAction 'Stop' }
	catch { }
	
	##*===============================================
	##* VARIABLE DECLARATION
	##*===============================================
	## Variables: Application
	[string]$appVendor = 'PwC IT'
	[string]$appName = 'Office 2013 KMS Activation'
	[string]$appVersion = '1.1'
	[string]$appArch = 'x86'
	[string]$appLang = 'EN'
	[string]$appRevision = '01'
	[string]$appScriptVersion = '1.0.0'
	[string]$appScriptDate = '04/01/2017'
	[string]$appScriptAuthor = 'PwC IT Global Desktop Team'
	[string]$KMSServer = "globalkms.pwcinternal.com"
	[string]$KMSPort = "1688"
	##*===============================================
	
	##* Do not modify section below
	#region DoNotModify
	
	## Variables: Exit Code
	[int32]$mainExitCode = 0
	
	## Variables: Script
	[string]$deployAppScriptFriendlyName = 'Deploy Application'
	[version]$deployAppScriptVersion = [version]'3.6.5'
	[string]$deployAppScriptDate = '08/17/2015'
	[hashtable]$deployAppScriptParameters = $psBoundParameters
	
	## Variables: Environment
	If (Test-Path -LiteralPath 'variable:HostInvocation') { $InvocationInfo = $HostInvocation }
	Else { $InvocationInfo = $MyInvocation }
	[string]$scriptDirectory = Split-Path -Path $InvocationInfo.MyCommand.Definition -Parent
	
	## Dot source the required App Deploy Toolkit Functions
	Try
	{
		[string]$moduleAppDeployToolkitMain = "$scriptDirectory\AppDeployToolkit\AppDeployToolkitMain.ps1"
		If (-not (Test-Path -LiteralPath $moduleAppDeployToolkitMain -PathType 'Leaf')) { Throw "Module does not exist at the specified location [$moduleAppDeployToolkitMain]." }
		If ($DisableLogging) { . $moduleAppDeployToolkitMain - DisableLogging }
		Else { . $moduleAppDeployToolkitMain }
	}
	Catch
	{
		If ($mainExitCode -eq 0) { [int32]$mainExitCode = 60008 }
		Write-Error -Message "Module [$moduleAppDeployToolkitMain] failed to load: `n$($_.Exception.Message)`n `n$($_.InvocationInfo.PositionMessage)" -ErrorAction 'Continue'
		## Exit the script, returning the exit code to SCCM
		If (Test-Path -LiteralPath 'variable:HostInvocation') { $script:ExitCode = $mainExitCode; Exit }
		Else { Exit $mainExitCode }
	}
	
	#endregion
	##* Do not modify section above
	##*===============================================
	##* END VARIABLE DECLARATION
	##*===============================================
	
	##* Show Progress Message (with the default message)
	Show-InstallationProgress
	
	##*===============================================
	##* Validating OS
	##*===============================================
	[string]$installPhase = 'Validating OS'
	Write-Log "Starting..."
	Write-Progress "OS Validation"
	Write-Log "Operating System name: $envOSName"
	if ([version]$envOS.version -ge [version]"10.0.0000")
	{
		Write-Log "Windows 10 OS detected" -Severity 3
		Show-DialogBox -Text "Unsupported OS detected" -Icon 'Stop'
		Exit-Script -ExitCode 69001
	}
	else
	{
		Write-Log -Message "Valid OS detected"
	}
	Write-Log "Complete"
	
	##*===============================================
	##* Licensing - Activation of Windows
	##*===============================================
	[string]$installPhase = 'Licensing'
	Write-Log "Starting..."
	Write-Progress "Licensing configuration"
	
	## Gather Windows activation details from slmgr.vbs
	$WindowsActivationDetails = Get-WindowsActivationDetails
	
	## If a MAK license key exists, replace it with the Global KMS license key for the relevant Operating System
	if ($WindowsActivationDetails.IsMAK -eq $true)
	{
		[string]$ShortVersionNumber = [string]$([version]$envOS.Version).Major + [string]'.' + [string]$([version]$envOS.Version).Minor
		Write-Log 'MAK license key found.'
		$WindowsKMSKeys = @{ "6.1" = "33PXH-7Y6KF-2VJC9-XBBR8-HVTHH"; "6.3" = "MHF9N-XY6XB-WVXMC-BTDCT-MKKG7" }
		
		Write-Log "Setting Global KMS key for $envOSName..."
		Execute-Process -Path "$envSystemRoot\System32\cscript.exe" -Parameters "$envSystemRoot\System32\slmgr.vbs -ipk $($WindowsKMSKeys.($ShortVersionNumber))" -WindowStyle Hidden
	}
	else
	{
		Write-Log 'No Operating System MAK license key found. Client must have KMS global license key installed from image.'
	}
	
	## If KMS server details do not match the ones specified in the KMSServer and KMSPort parameters, change them
	if (!($WindowsActivationDetails.KMSHost -like "$KMSServer`:$KMSPort"))`
	{
		Write-Log 'KMS server details do not exist or are incorrect, setting KMS Server details...'
		Execute-Process -Path "$envSystemRoot\System32\cscript.exe" -Parameters "$envSystemRoot\System32\slmgr.vbs -skms $KMSServer`:$KMSPort" -WindowStyle Hidden
	}
	else
	{
		Write-Log "KMS server details for $envOSName are already set as $KMSServer`:$KMSPort`."
	}
	
	##*===============================================
	##* Licensing - Activation of Office
	##*===============================================
	
	##* Reconfigure Office 2013 for KMS Activation
	if (Get-InstalledApplication -Name "Microsoft Office Professional Plus 2013")
	{
		$ProductKMSKeys = @{ "Office" = "YC7DK-G2NP3-2QQC3-J6H88-GVGXT"; "Visio" = "C2FG9-N6J68-H8BTJ-BW3QX-RM3B3"; "Project" = "FN8TT-7WMH6-2D4X9-M337T-2342K" }
		$OfficePath = (Get-RegistryKey -Key "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\15.0\Common\InstallRoot" -Value Path)
        $OSPPVbs = '"' + $OfficePath + 'OSPP.VBS' + '"'
		Write-Log "Office 2013 path: $OfficePath"
		$OfficeActivationDetails = Get-OfficeActivationDetails -officePath $OfficePath
		foreach ($OfficeProduct in $OfficeActivationDetails)
		{
			If ($OfficeProduct.IsMAK -like "True")
			{
				## Unpublish the existing MAK Key
				Write-Log "Remove $($OfficeProduct.Product) MAK: $($OfficeProduct.Last5)"
				Execute-Process -Path "$envSystemRoot\System32\cscript.exe" -Parameters "$OSPPVbs /unpkey:$($OfficeProduct.Last5)" -WindowStyle Hidden
				#Publish the Global Volume License Key for the product
				Write-Log "Add GVLK for $($OfficeProduct.Product)"
				Execute-Process -Path "$envSystemRoot\System32\cscript.exe" -Parameters "$OSPPVbs /inpkey:$($ProductKMSKeys.Get_Item($($OfficeProduct.Product)))" -WindowStyle Hidden
				$enableActivation = $true
			}
			else
			{
				Write-Log "$($OfficeProduct.Product) does not have an existing MAK key."
			}
			If (!($OfficeProduct.KMSHost -like "$KMSServer`:$KMSPort"))
			{
				Write-Log "Set Global KMS Server details"
				Execute-Process -Path "$envSystemRoot\System32\cscript.exe" -Parameters "$OSPPVbs /sethst:$KMSServer" -WindowStyle Hidden
				Execute-Process -Path "$envSystemRoot\System32\cscript.exe" -Parameters "$OSPPVbs /setprt:$KMSPort" -WindowStyle Hidden
				$enableActivation = $true
			}
			else
			{
				Write-Log "KMS server details for $($OfficeProduct.Product) are already set as $KMSServer`:$KMSPort`."
			}
		}

        # Activate products
		Write-Log 'Activating Windows via KMS...'
		Execute-Process -Path "$envSystemRoot\System32\cscript.exe" -Parameters "$envSystemRoot\System32\slmgr.vbs -ato" -WindowStyle Hidden -ContinueOnError $true
		Write-Log "Activate Office"
		Execute-Process -Path "$envSystemRoot\System32\cscript.exe" -Parameters "$OSPPVbs /act" -WindowStyle Hidden -ContinueOnError $true
	}
	
	Write-Log "Complete"
	
	## Call the Exit-Script function to perform final cleanup operations
	Exit-Script -ExitCode $mainExitCode
}
catch
{
	[int32]$mainExitCode = 60001
	[string]$mainErrorMessage = "$(Resolve-Error)"
	Write-Log -Message $mainErrorMessage -Severity 3 -Source $deployAppScriptFriendlyName
	Show-DialogBox -Text $mainErrorMessage -Icon 'Stop'
	Exit-Script -ExitCode $mainExitCode
}