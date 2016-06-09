#========================================================================
# Created with: SAPIEN Technologies, Inc., PowerShell Studio 2012 v3.1.35
# Created on: 6/30/2015 2:39 PM
# Created by: Craig Woodford (craigw@umn.edu) Jeff Bolduan (jbolduan@umn.edu)
# Organization: University of Minnesota - OIT
# Filename: Create-SharedFolder.ps1   
#========================================================================

<#
	.SYNOPSIS
		This script creates shared folders in the OIT managed environment.
	.DESCRIPTION
		Create-SharedFolder creates a folder and groups for modify and read-only rights.  It applies modify and read-only
		rights respectively on the folder.  It also creates a group management group and gives that group the rights to manage
		group membership for the modify and read-only groups.
	.PARAMETER folderName
		The name of the folder. This should be a legal folder name (no illegal charcters or trailing spaces).
	.PARAMETER folderPath
		The full path to folder in which you want to create the new folder.
	.PARAMETER unit
		This should be a unit managed by OIT existing under the ad.umn.edu\clients\units\ OU.
	.PARAMETER subUnit
		This is an optional sub-unit.
	.PARAMETER outConsole
		Determines if messages are logged to the console as well as the log file.  Default is $false.
	.PARAMETER logPath
		The full path of the folder where the log file will be created. Default is the current directory.
	.EXAMPLE
		Create-SharedFolder.ps1 -folderName "TestName" -folderPath "\\files.umn.edu\OIT\Shared" -unit "OIT" -subUnit "CMD"
			Creates the folder: \\files.umn.edu\OIT\Shared\TestName
			Creates the groups: OIT-FS-CMD-TestName-M
								OIT-FS-CMD-TestName-R
								OIT-FS-MGRS-CMD-TestName
			Applies modify rights to OIT-FS-CMD-TestName-M and read-only rights to OIT-FS-CMD-TestName-R on the newly created folder.
			Sets the OIT-FS-MGRS-CMD-TestName as the manager of both the OIT-FS-CMD-TestName-M and OIT-FS-CMD-TestName-R groups and
			checks the box allowing the manager to edit group membership.
			Logs all actions to a log file created in the directory in which the script is run.
#>

# Parameter setup
param (
	[Parameter(Mandatory=$true,
		HelpMessage="Enter a folder name")]
	[string]
	$folderName
	,
	[Parameter(Mandatory=$true,
		HelpMessage="Enter a valid destination directory in the form of \\server\share\folder")]
	[string]
	$folderPath
	,
	[Parameter(Mandatory=$true,
		HelpMessage="Enter a valid unit prefix (without a -)")]
	[string]
	$unit
	,
	[Parameter(Mandatory=$false,
		HelpMessage="Enter a valid sub-unit accronym")]
	[string]
	$subUnit = ""
	,
	[Parameter(Mandatory=$false,
		HelpMessage="Enter a true boolean if you want to display log contents to the console")]
	[bool]
	$outConsole = $false
	,
	[Parameter(Mandatory=$false,
		HelpMessage="Enter a valid directory for the log file")]
	[string]
	$logPath = (Get-Location).Path
) # endparam

Import-Module ActiveDirectory

#region Functions

Function Write-Log {
	<#
		.SYNOPSIS
			This function is used to pass messages to a ScriptLog.  It can also be leveraged for other purposes if more complex logging is required.
		.DESCRIPTION
			Write-Log function is setup to write to a log file in a format that can easily be read using CMTrace.exe. Variables are setup to adjust the output.
		.PARAMETER Message
			The message you want to pass to the log.
		.PARAMETER Path
			The full path to the script log that you want to write to.
		.PARAMETER Severity
			Manual indicator (highlighting) that the message being written to the log is of concern. 1 - No Concern (Default), 2 - Warning (yellow), 3 - Error (red).
		.PARAMETER Component
			Provide a non null string to explain what is being worked on.
		.PARAMETER Context
			Provide a non null string to explain why.
		.PARAMETER Thread
			Provide a optional thread number.
		.PARAMETER Source
			What was the root cause or action.
		.PARAMETER Console
			Adjusts whether output is also directed to the console window.
		.NOTES
			Name: Write-Log
			Author: Aaron Miller
			LASTEDIT: 01/23/2013 10:09:00
		.EXAMPLE
			Write-Log -Message $exceptionMsg -Path $ScriptLog -Severity 3
			Writes the content of $exceptionMsg to the file at $ScriptLog and marks it as an error highlighted in red
	#>

	PARAM(
		[Parameter(Mandatory=$True)][String]$Message,
		[Parameter(Mandatory=$False)][String]$Path = "$env:TEMP\CMTrace.Log",
		[Parameter(Mandatory=$False)][int]$Severity = 1,
		[Parameter(Mandatory=$False)][string]$Component = " ",
		[Parameter(Mandatory=$False)][string]$Context = " ",
		[Parameter(Mandatory=$False)][string]$Thread = "1",
		[Parameter(Mandatory=$False)][string]$Source = "",
		[Parameter(Mandatory=$False)][switch]$Console
	)
			
	# Setup the log message
	
		$time = Get-Date -Format "HH:mm:ss.fff"
		$date = Get-Date -Format "MM-dd-yyyy"
		$LogMsg = '<![LOG['+$Message+']LOG]!><time="'+$time+'+000" date="'+$date+'" component="'+$Component+'" context="'+$Context+'" type="'+$Severity+'" thread="'+$Thread+'" file="'+$Source+'">'
			
	# Write out the log file using the ComObject Scripting.FilesystemObject
	
		$ForAppending = 8
		$oFSO = New-Object -ComObject scripting.filesystemobject
		$oFile = $oFSO.OpenTextFile($Path, $ForAppending, $True)
		$oFile.WriteLine($LogMsg)
		$oFile.Close()
		Remove-Variable oFSO
		Remove-Variable oFile
		
	# Write to the console if $Console is set to True
	
		if ($Console -eq $True) {Write-Host $Message}
		
} # End Write-Log

Function validGroup {
<#
	.SYNOPSIS
		Returns true if a group exists.
	
	.PARAMETER groupName
		The name of the group to check.
#>
	Param (
		[Parameter(Mandatory=$True)]$groupName
	)

	try {
		Get-ADGroup $groupName | Out-Null
		return $true
	}
	catch {
		return $false
	}
	
} # End validGroup

Function validOU {
<#
	.SYNOPSIS
		Returns true if an OU exists.
	
	.PARAMETER distinguishedName
		The distinguished name of the OU to check.
#>
	Param (
		[Parameter(Mandatory=$True)]$distinguishedName
	)

	try {
		Get-ADOrganizationalUnit $distinguishedName | Out-Null
		return $true
	}
	catch {
		return $false
	}
	
} # End validOU

Function validADObject {
<#
	.SYNOPSIS
		Returns true if an AD object is valid on a specific domain controller.
	
	.PARAMETER objName
		The AD object to check.
	
	.PARAMETER Server
		The domain controller to check against.
#>
	Param (
		[Parameter(Mandatory=$True)]
		[string]$objName
		,
		[Parameter(Mandatory=$true)]
		[string]$Server
	)

	try {
		Get-ADObject -Filter {Name -eq $objName} -Server $server | Out-Null
		return $true
	}
	catch {
		return $false
	}
	
} # End validADObject

Function checkDCs {
<#
	.SYNOPSIS
		Check's if an AD object exists on each domain controller within a domain.
	
	.DESCRIPTION
		This can be used to verify that a newly created object has replicated to each
		domain controller before proceeding with other actions.
	
	.PARAMETER ADObject
		The AD object to check.
	
	.PARAMETER maxWait
		The maximum time in seconds to wait.
#>
	
	Param (
	[Parameter(Mandatory=$true)]
	[string]$ADObject
	,
	[Parameter(Mandatory=$False)]
	[int]$maxWait = 30
	)
	
	$dcList = New-Object System.Collections.ArrayList
	
	$count = 1
	
	try {
		# Get the list of domain controllers to check
		Get-ADDomainController -Filter * | % {$dcList.Add($_.HostName)}
		
		while (($count -le $maxWait) -and ($dcList)) {
		
			$goodServers = New-Object System.Collections.ArrayList
			
			# Walk through each DC and identify the DC's that see the AD Object
			$dcList | % {
				
				if(validADObject -objName $ADObject -Server $_) {
					$goodServers.Add($_)
				}
			}
			
			# Remove the DC's that already see the AD object from the list of DC's to check
			$goodServers | % {$dcList.Remove($_)}
			
			Start-Sleep -Seconds 1
			
			$count += 1
		}
		
		if(!($dcList)) {
			Return $true
		}
		else {
			Return $false
		}
	}
	catch {
		throw
	}
} # End checkDCs

function Set-GroupManager {
	<#
		.Synopsis
		   Sets manager property on AD group and grants change membership rights.
		.DESCRIPTION
		   Sets manager property on AD group and grants change membership rights.
		   This is done by manipulating properties directly on the DirectoryEntry object
		   obtained with ADSI. This sets the managedBy property and adds an ACE to the DACL
		   allowing said manager to modify group membership.
		   Taken from: https://www.bazaarbytes.com/blog/setting-ad-group-managers-with-powershell/
		.EXAMPLE
		   Set-GroupManager -ManagerDN "CN=some manager,OU=All Users,DC=Initech,DC=com" -GroupDN "CN=TPS Reports Dir,OU=All Groups,DC=Initech,DC=com"
		.EXAMPLE
		   (Get-AdGroup -Filter {Name -like "sharehost - *"}).DistinguishedName | % {Set-GroupManager "CN=some manager,OU=All Users,DC=Initech,DC=com" $_}
	#>
    param (
        [Parameter(Mandatory=$true, ValueFromPipeline=$false, ValueFromPipelinebyPropertyName=$True, Position=0)]
        [string]$ManagerDN,
        [Parameter(Mandatory=$true, ValueFromPipeline=$false, ValueFromPipelinebyPropertyName=$True, Position=1)]
        [string]$GroupDN,
		[Parameter(Mandatory=$true, ValueFromPipeline=$false, ValueFromPipelinebyPropertyName=$True, Position=1)]
        [string]$targetDC
        )
    
    try {
 
        $mgr = [ADSI]"LDAP://$targetDC/$ManagerDN";
        $identityRef = (Get-ADGroup -Filter {DistinguishedName -like $ManagerDN} -Server $targetDC).SID.Value
        $sid = New-Object System.Security.Principal.SecurityIdentifier ($identityRef);
 
        $adRule = New-Object System.DirectoryServices.ActiveDirectoryAccessRule ($sid, `
                    [System.DirectoryServices.ActiveDirectoryRights]::WriteProperty, `
                    [System.Security.AccessControl.AccessControlType]::Allow, `
                    [Guid]"bf9679c0-0de6-11d0-a285-00aa003049e2");
 
        $grp = [ADSI]"LDAP://$targetDC/$GroupDN";
 
        # Taken from here: http://blogs.msdn.com/b/dsadsi/archive/2013/07/09/setting-active-directory-object-permissions-using-powershell-and-system-directoryservices.aspx
        [System.DirectoryServices.DirectoryEntryConfiguration]$SecOptions = $grp.get_Options();
        $SecOptions.SecurityMasks = [System.DirectoryServices.SecurityMasks]'Dacl'
                
        $grp.get_ObjectSecurity().AddAccessRule($adRule);
        $grp.CommitChanges();
    }
    catch {
        throw
    }
} # End Set-GroupManager

Function Apply-SharedFolderPerms {
<#
	.SYNOPSIS
		Apply modify and read permissions to groups on a folder. 
	
	.DESCRIPTION
		If there is an error referencing 'SetAccessRule' then keep trying until it works or the maxWait 
		time is reached. Errors referencing 'SetAccessRule' when applying permissions almost always mean 
		that the file share can not yet enumerate one or more of the group names.  This can happen when 
		groups are created via scripts.  The 'SetAccessRule' error message may be different in different 
		languages.
	
	.PARAMETER pathName
		The path of the folder to apply permissions to.
	
	.PARAMETER modifyName
		The name of the group that gets modify rights.
	
	.PARAMETER readName
		The name of the group that gets read rights.
	
	.PARAMETER domainName
		The domain name with a '\' at the end. (for example: ad.umn.edu\)
	
	.PARAMETER maxWait
		The maximum time in seconds to wait.
#>	
	Param (
	[Parameter(Mandatory=$true)]
	[string]$pathName
	,
	[Parameter(Mandatory=$true)]
	[string]$modifyName
	,
	[Parameter(Mandatory=$true)]
	[string]$readName
	,
	[Parameter(Mandatory=$true)]
	[string]$domainName
	,
	[Parameter(Mandatory=$False)]
	[int]$maxWait = 30
	)	
	
	$count = 1
	$shareStatus = $false
	
	while (($count -le $maxWait) -and (!($shareStatus))) {
		
		try {
			# Get the ACL of the folder
			$acl =  [System.IO.Directory]::GetAccessControl($pathName)
			
			# Create the permission
			$modPermission = "$domainName$modifyName","Modify","ContainerInherit,ObjectInherit","None","Allow"
			$readPermission = "$domainName$readName","ReadAndExecute","ContainerInherit,ObjectInherit","None","Allow"

			# Create the access rule
			$modifyAccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule $modPermission
			$readAccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule $readPermission

			# Set the ACL
			$acl.SetAccessRule($modifyAccessRule)
			$acl.SetAccessRule($readAccessRule)
			
			# Apply the ACL to the folder
			[System.IO.Directory]::SetAccessControl($pathName, $acl)
			
			$shareStatus = $true
			$count += $maxWait
		}
		catch {
			if ($_.Exception.Message -like "*SetAccessRule*") {
				# Catch an error referencing 'SetAccessRule'.  This string may be different in different languages.
				# Sleep 1 second then increment the counter to return to the while loop.
				Start-Sleep -Seconds 1
				$count += 1
			}
			else {
				# This is a different error just throw it and exit the function
				throw $_	
			}
		}
	}
	
	if ($shareStatus) {
		Return $true
	}
	else {
		Return $false	
	}
} # End Apply-SharedFolderPerms

#endregion

#region variables

$fullPath = $folderPath + '\' + $folderName

$dateTime = Get-Date -uformat %Y%m%d-%H%M%S

$logName = $logPath + '\' + "CMD-CSF-" + $unit + "-" + $folderName + "_" + $dateTime + ".log"

$baseOUPath = "OU=Units,OU=Clients,DC=ad,DC=umn,DC=edu"

$unitOUPath = "OU=$unit,$baseOUPath"

$fsOUPath = "OU=Fileshare,OU=Groups,$unitOUPath"

$mgOUPath = "OU=Management,OU=Groups,$unitOUPath"

if ($subUnit -ne "") {
	$subUnitName = "$subUnit-"
}
else {
	$subUnitName = ""	
}

$modifyGroupName = "$unit-FS-$subUnitName$folderName-M"

$readGroupName = "$unit-FS-$subUnitName$folderName-R"

$manageGroupName = "$unit-FS-MGRS-$subUnitName$folderName"

$modifyDesc = "Modify to: $fullPath"

$readDesc = "Read only to: $fullPath"

$manageDesc = "Controls access to: $fullPath"

$domainController = (Get-ADDomainController).HostName

# The AD domain name
$domainName = "ad.umn.edu\"

#endregion

#region main

# Test if $logPath exists
if (!(Test-Path $logPath)) {
	
	try {
		# Create the $logPath if it doesn't exist
		New-Item -Path $logPath -Type Directory -Force
		Write-Log -Message "Created $logPath" -Path $logName -Console:$outConsole
	}
	catch {
		# We couldn't create the $logPath, we need to exist the script.
		Throw "Error! Script exiting: unable to create $logPath!!!"
		Exit 666
	}
}

try {
	# Log initial variables
	Write-Log -Message "createSharedFolder started at: $dateTime with the following variables:" -Path $logName -Console:$outConsole
	Write-Log -Message "folderName: $folderName" -Path $logName -Console:$outConsole
	Write-Log -Message "folderPath: $folderPath" -Path $logName -Console:$outConsole
	Write-Log -Message "unit: $unit" -Path $logName -Console:$outConsole
	Write-Log -Message "subUnit: $subUnit" -Path $logName -Console:$outConsole
	Write-Log -Message "outConsole: $outConsole" -Path $logName -Console:$outConsole
	Write-Log -Message "logPath: $logPath" -Path $logName -Console:$outConsole
}
catch {
	# We couldn't write to a logfile, we need to exit the script.
	#throw "Error! Script exiting: unable to write to $logName!!!"
	Write-Host "error with script:"
	throw $_
	Exit 666
}

#region validation

# Validate all variables
try {

	if (!(Test-Path $folderPath)) {
		
		Write-Log -Message "Error: $folderPath does not exist." -Path $logName -Console:$outConsole
		Throw "Exiting due to error with $folderPath"	
	}
	
	if (Test-Path $fullPath) {
		
		Write-Log -Message "Error: $fullPath exists." -Path $logName -Console:$outConsole
		Throw "Exiting due to $fullPath already existing."	
	}
	
	if (!(validOU -distinguishedName $unitOUPath)) {
	
		Write-Log -Message "Error: $unitOUPath does not exist." -Path $logName -Console:$outConsole
		Throw "Exiting due to error with $unitOUPath"
	}
	
	if (!(validOU -distinguishedName $fsOUPath)) {
	
		Write-Log -Message "Error: $fsOUPath does not exist." -Path $logName -Console:$outConsole
		Throw "Exiting due to error with $fsOUPath"
	}
	
	if (!(validOU -distinguishedName $mgOUPath)) {
	
		Write-Log -Message "Error: $mgOUPath does not exist." -Path $logName -Console:$outConsole
		Throw "Exiting due to error with $mgOUPath"
	}
	
	if (validGroup -groupName $modifyGroupName) {
	
		Write-Log -Message "Error: $modifyGroupName already exists." -Path $logName -Console:$outConsole
		Throw "Exiting due to $modifyGroupName already existing."
	}
	
	if (validGroup -groupName $readGroupName) {
	
		Write-Log -Message "Error: $readGroupName already exists." -Path $logName -Console:$outConsole
		Throw "Exiting due to $readGroupName already existing."
	}
	
	if (validGroup -groupName $manageGroupName) {
	
		Write-Log -Message "Error: $manageGroupName already exists." -Path $logName -Console:$outConsole
		Throw "Exiting due to $manageGroupName already existing."
	}
}
catch {
	# A variable wasn't validated, we need to exit the script.
	Write-Log -Message "Error! Script exiting, error message to follow:" -Path $logName -Console:$outConsole
	Write-Log -Message $_.Exception.ToString() -Path $logName -Console:$outConsole
	Exit 666
}
#endregion

try {
	# Create the folder
	New-Item -Path $folderPath -Name $folderName -Type Directory -Force | Out-Null
	Write-Log -Message "Created $fullPath" -Path $logName -Console:$outConsole
}
catch {
	# We couldn't create the new folder, we need to exit the script.
	Write-Log -Message "Problem creating: $fullPath" -Path $logName -Console:$outConsole
	Write-Log -Message "Error! Script exiting, error message to follow:" -Path $logName -Console:$outConsole
	Write-Log -Message $_.Exception.ToString() -Path $logName -Console:$outConsole
	Exit 666
}

try {
	# Create AD Groups
	New-ADGroup -Name $manageGroupName -Path $mgOUPath -Description $manageDesc -GroupScope Global -Server $domainController
	Write-Log -Message "Created group to control access: $manageGroupName" -Path $logName -Console:$outConsole
	
	New-ADGroup -Name $modifyGroupName -Path $fsOUPath -Description $modifyDesc -GroupScope Global -Server $domainController -ManagedBy $manageGroupName
	Write-Log -Message "Created group for modify access: $modifyGroupName" -Path $logName -Console:$outConsole
	
	New-ADGroup -Name $readGroupName -Path $fsOUPath -Description $readDesc -GroupScope Global -Server $domainController -ManagedBy $manageGroupName
	Write-Log -Message "Created group for read access: $readGroupName" -Path $logName -Console:$outConsole
	
	$managerDName = (Get-ADGroup $manageGroupName -Server $domainController).DistinguishedName
	$modifyDName = (Get-ADGroup $modifyGroupName -Server $domainController).DistinguishedName
	$readDName = (Get-ADGroup $readGroupName -Server $domainController).DistinguishedName
	
	# Check the Manager can update membership list check-box for the modify and read groups
	Set-GroupManager -ManagerDN $managerDName -GroupDN $modifyDName -targetDC $domainController
	Write-Log -Message "Checked the Manager can update membership list check-box for: $modifyGroupName" -Path $logName -Console:$outConsole
	Set-GroupManager -ManagerDN $managerDName -GroupDN $readDName -targetDC $domainController
	Write-Log -Message "Checked the Manager can update membership list check-box for: $readGroupName" -Path $logName -Console:$outConsole
	
	# Verify that all created groups have propagated to all Domain Controllers
	if (checkDCs -ADObject $manageGroupName) {
		Write-Log -Message "Verified that $manageGroupName has propagated to all Domain Controllers" -Path $logName -Console:$outConsole
	}
	else {
		Write-Log -Message "Unable to verify that $manageGroupName has propagated to all Domain Controllers... continuing anyways." -Path $logName -Console:$outConsole
	}
	
	if (checkDCs -ADObject $modifyGroupName) {
		Write-Log -Message "Verified that $modifyGroupName has propagated to all Domain Controllers" -Path $logName -Console:$outConsole
	}
	else {
		Write-Log -Message "Unable to verify that $modifyGroupName has propagated to all Domain Controllers... continuing anyways." -Path $logName -Console:$outConsole
	}
	
	if (checkDCs -ADObject $readGroupName) {
		Write-Log -Message "Verified that $readGroupName has propagated to all Domain Controllers" -Path $logName -Console:$outConsole
	}
	else {
		Write-Log -Message "Unable to verify that $readGroupName has propagated to all Domain Controllers... continuing anyways." -Path $logName -Console:$outConsole
	}
	
	# Apply permissions to the folder
	if (Apply-SharedFolderPerms -domainName $domainName -modifyName $modifyGroupName -readName $readGroupName -pathName $fullPath) {
		Write-Log -Message "ACLs successfully applied to the folder." -Path $logName -Console:$outConsole
	}
	else {
		Write-Log -Message "Error writing ACLs to $fullPath" -Path $logName -Console:$outConsole
		Throw "Error writing ACLs.  Most likely group object creation did not propagate to the file share."
	}
}
catch {
	# There was a problem with the script, we need to exit.
	Write-Log -Message "Error! Script exiting, error message to follow:" -Path $logName -Console:$outConsole
	Write-Log -Message $_.Exception.ToString() -Path $logName -Console:$outConsole
	Exit 666
}

#endregion	