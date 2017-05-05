<#
.SYNOPSIS
	Module of re-usable functions

.DESCRIPTION
	Util.psm1 is a collection of functions that can be leveraged by other scripts
	To utilise this module simply reference it from your script:
		## Get the Scripts Directory
		$scriptDir = Split-Path $MyInvocation.MyCommand.Path
		## Import Util Module
		Import-Module "$scriptDir\Include\Util.psm1" -force

.NOTES
	Original Author: 	Jacob Hodges (Technology Effect)
	Creation Date: 		2010/11/30
	Maintained by:		David Bubb (Technology Effect)
    Last Change Author: Jesse Harris (University of the Sunshine Coast)
	Last Change:		2015/08/06
    
	
	*** EACH FUNCTION MUST INCLUDE COMMENTS/EXAMPLES OF USAGE

.FUNCTIONS
	*** PLEASE MAINTAIN THIS LIST OF FUNCTIONS AS CHANGES ARE MADE

	Write-Entry				: 	Write a line to a log file (in format that is best viewed with Trace32.exe)
	Get-Settings			:	Read and store appropriately the contents of a XML file
	Copy-LogFile			:	Copy $global:logFile to the specified location
	Set-PinnedItem			:	Pin (or UnPin) an icon to (from) the Start Menu or Task Bar
	New-LocalAccount		:	Create a local user account
	Add-AccountToLocalGroup	:	Add a user to a local group
	Disable-LocalAccount	:	Disable a local user account
	Find-ADObject			:	Search for an Active Directory object
	Move-ADObject			:	Move an Active Directory object to a different OU
	Get-Gateway				:	Obtain the Gateway IP address of the local machine
	Get-Laptop				:	Determine if local machine is a laptop device
	Create-Directory		:	Create a folder
	Rename-Directory		:	Rename a folder
	Move-Directory			:	Move a folder
	Configure-ACL			:	Set the ACL (NTFS permissions) for a specified folder or file
	Create-Shortcut			:	Create a shortcut
	Disable-Service			:	Disable a Windows Service - SUPERCEDED by Change-Service-Startup
	Write-Registry			:	Set a value in the registry
	Disable-SchedTask		:	Disable a Scheduled Task
	Halt-Service			:	Stop a Windows Service
	Delete-RegistryValue	:	Delete a value from the registry
	Delete-Files			:	Delete files from a specified folder based on a specified file pattern
	Change-Service-Startup	:	Change the startup type for a Windows service
	Show-MessageBox			:	Displays a MessageBox using Windows WinForms
	Clear-SCCMCache         :   Clear SCCM Cache
    Get-SCCMCacheInfo       :   Retrieve SCCM Cache Object
#>

#Set some Globel Variables
$global:LogTypeInfo = 1
$global:LogTypeWarning = 2
$global:LogTypeError = 3
$global:LogTypeVerbose = 4
$global:LogTypeDeprecated = 5
$global:Success = 0
$global:Failure = 1


function Write-Entry {
    <#
        .DESCRIPTION
            Write a line to a log file (in format that is best viewed with Trace32.exe)
            
        .PARAMETER logMsg (Mandatory)
            String;	Text to write to log
		
        .PARAMETER msgType
            Int;	Info=1, Warning=2, Error=3, Verbose=4, Deprecated=5
					[Utilise global variables LogTypeInfo (default), LogTypeWarning, LogTypeError, LogTypeVerbose or LogTypeDeprecated]
		
        .OUTPUT/RETURN
			Output = Debuggin message
		
		.EXAMPLE
            Write-Entry "Action was successful"
			Write-Entry "Error: $_" $global:LogTypeError
		
		.NOTES
			Author:		Jacob Hodges
			Created:	2010/11/30
	#>
	
    [CmdletBinding()]
    PARAM
    (
        [Parameter(Position=1, Mandatory=$true)] $logMsg,
        [Parameter(Position=2)] $msgType = $global:LogTypeInfo
    )

	#Populate the variables to log
	$time = [DateTime]::Now.ToString("HH:mm:ss.fff+000");
	$date = [DateTime]::Now.ToString("MM-dd-yyyy");
	$component = $myInvocation.ScriptName | Split-Path -leaf
	$file = $myInvocation.ScriptName
	
	$tempMsg = [String]::Format("<![LOG[{0}]LOG]!><time=`"{1}`" date=`"{2}`" component=`"{3}`" context=`"`" type=`"{4}`" thread=`"`" file=`"{5}`">",$logMsg, $time, $date, $component, $msgType, $file)
    
	if($debug)
	{
		Write-Host $logMsg
	}
	
	$tempMsg | Out-File -encoding ASCII -Append -FilePath $global:logFile 
}

function Get-Settings {
    <#
        .DESCRIPTION
            Read and store appropriately the contents of a XML file
            
        .PARAMETER settingsFile (Mandatory)
            String;	File to read (including path)
		
        .OUTPUT/RETURN
			Return = XML object

		.EXAMPLE
            $global:contentXML = Get-Settings $controlXML
		
		.NOTES
			Author:		Jacob Hodges
			Created:	2010/11/30
	#>

    [CmdletBinding()]
    PARAM
    (
        [Parameter(Position=1, Mandatory=$true)] $settingsFile
    )

    # Get the file contents
    $contents = Get-Content $settingsFile

    # Turn it into an XML object
    $settings = [xml] $contents

    # Return the XML object
    $settings 
}

function Copy-LogFile {
    <#
        .DESCRIPTION
            Copy $global:logFile to the specified location
            
        .PARAMETER Path
            The path to copy the log file to

        .OUTPUT/RETURN
			Nothing
			
		.EXAMPLE
            Copy-LogFile C:\Support
		
		.NOTES
			Author:		Jacob Hodges
			Created:	2010/11/30
	#>
    
    [CmdletBinding()]
    PARAM
    (
        [Parameter(Position=1, Mandatory=$true)] [string]$Path
    )
    
    #Verify that the path exists
    if(Test-Path $Path)
    {
        #Path Exists copy the file
        Copy-Item $global:logFile $Path -ErrorAction SilentlyContinue
    }

}

function Set-PinnedItem {
    <#
        .DESCRIPTION
            Pins an item to the Start Menu or Task Bar
            
        .PARAMETER path
            The Path of the item to pin to the task par
		
        .PARAMETER StartMenu
            [switch] Creates the Pinned Item in the Start Menu instead of the taskbar
		
		.PARAMETER Unpin
			[switch] Unpins the item rather than pinning it
		
        .OUTPUT/RETURN
			Nothing		
		
        .EXAMPLE
            Set-PinnedItem C:\Windows\Explorer.exe
		
		.NOTES
			Author:		Jacob Hodges
			Created:	2010/10/07
	#>

	[CmdletBinding()]
	param (
		[parameter(Mandatory=$true)]
		[string]$Path,
		[switch]$StartMenu,
		[switch]$Unpin
	)

	Process {
		

		$app = New-Object -ComObject "Shell.Application"
		$folder = $app.Namespace((Split-Path $Path -Parent))
		$folderItem = $folder.ParseName((Split-Path $Path -Leaf))
		
		$PinOrUnPin = "Pin to "
		if($Unpin){$PinOrUnPin = "Unpin from "}
		
		$verbString = $PinOrUnPin + "Taskbar"
		if($StartMenu){$verbString = $PinOrUnPin + "Start Menu"}
			
		$folderItem.Verbs() | ?{$_.name.Replace("&","") -eq $verbString} | %{$_.DoIt()}

	}
	
}

function New-LocalAccount {
    <#
        .DESCRIPTION
            Create a local user account
            
        .PARAMETER Username
            Account name to create
		
        .PARAMETER Password
            Password to set for the account
		
		.PARAMETER FullName
			Full Name of the user
		
		.PARAMETER Description
			Description for the account
		
		.PARAMETER PasswordNeverExpires
			[switch] Enable non-expiring password for the account
		
        .OUTPUT/RETURN
			Nothing		
		
        .EXAMPLE
            New-LocalAccount -Username "smithf" -Password "Gu355Me" -FullName "Fred Smith" -Description "Fred's account" -PasswordNeverExpires
		
		.NOTES
			Author:		Jacob Hodges
			Created:	2010/10/07
	#>

    [CmdletBinding()]
    PARAM
    (
        [Parameter(Position=1, Mandatory=$true)] [string]$Username,
        [Parameter(Position=2, Mandatory=$true)] [string]$Password,
		[Parameter(Position=3, Mandatory=$true)] [string]$FullName,
		[Parameter(Position=4)] $Description,
		[Parameter(Position=5)] [switch]$PasswordNeverExpires
    )
	
	#Get the local WinNT Provider
	$winNT = [adsi]"WinNT://$($Env:COMPUTERNAME)"
	$user = $winNT.Create("User",$Username)
	$user.SetPassword($Password)
	$user.SetInfo()
	$user.FullName = $FullName
	$user.Description = $Description
	
	if($PasswordNeverExpires)
	{
		$user.Put("UserFlags",0x10000)
	}
	
	$user.SetInfo()

}

function Add-AccountToLocalGroup {
    <#
        .DESCRIPTION
            Add an account to an existing local machine group
            
        .PARAMETER Username
            Username to add to the group
		
        .PARAMETER Domain
            Domain that the user exists in (defaults to local host)
		
		.PARAMETER Group
			Local group to add user to
		
        .OUTPUT/RETURN
			Nothing		
		
        .EXAMPLE
            Add-AccountToLocalGroup -username "smithf" -group "Backup Operators"
		
		.NOTES
			Author:		Jacob Hodges
			Created:	2010/10/07
	#>

    [CmdletBinding()]
    PARAM
    (
        [Parameter(Position=1, Mandatory=$true)] [string]$Username,
		[Parameter(Position=2)] [string]$Domain = $Env:COMPUTERNAME,
        [Parameter(Position=3, Mandatory=$true)] [string]$Group
    )
	
	#Get the local WinNT Provider
	$computer = [adsi]"WinNT://$($Env:COMPUTERNAME),computer"
	$localGroup = $computer.Children.Find($Group)
	
	$localGroup.Add("WinNT://$Domain/$Username")
	
}

function Disable-LocalAccount {
    <#
        .DESCRIPTION
            Disable a local user account
            
        .PARAMETER Username
            Account to disable
		
        .OUTPUT/RETURN
			Nothing		
		
        .EXAMPLE
            Disable-LocalAccount -username "smithf"
		
		.NOTES
			Author:		Jacob Hodges
			Created:	2010/10/07
	#>

    [CmdletBinding()]
    PARAM
    (
        [Parameter(Position=1, Mandatory=$true)] [string]$Username
    )
	
	#Get the local WinNT Provider
	$computer = [adsi]"WinNT://$($Env:COMPUTERNAME),computer"
	$localAccount = $computer.Children.Find($Username)
	
	$localAccount.AccountDisabled = $true
	$localAccount.SetInfo()
	
}

function Find-ADObject {
    <#
        .DESCRIPTION
            Find an Active Directory object
            
        .PARAMETER Name
            Name (for computers) or samid (for all other objects) to search for
		
        .PARAMETER Type
            Object type (default is user)
		
        .OUTPUT/RETURN
			Return = System.DirectoryServices object
		
        .EXAMPLE
            Find-ADObject -name "jonest"
			Find-ADObject -name "hp7500-01" -type "computer"
			Find-ADObject -name "HR" -type "group"
		
		.NOTES
			Author:		David Bubb
			Created:	2010/13/12
	#>

    [CmdletBinding()]
    PARAM
    (
        [Parameter(Position=1, Mandatory=$true)]  [string]$Name,
		[Parameter(Position=2, Mandatory=$false)]  [string]$Type
    )
	
	$root = [ADSI]""
	$searcher = new-object System.DirectoryServices.DirectorySearcher($root)
	if ($Type -eq $null) {
		$searcher.filter = "(sAMAccountName=$Name)"
	} elseif ($Type.ToLower() -eq "computer") {
		$searcher.filter = "(&(objectClass=$Type)(Name=$Name))"
	} else {
		$searcher.filter = "(&(objectClass=$Type)(sAMAccountName=$Name))"
	}
	$DN = $searcher.findall()
      
	if (-not($DN.count -eq 1))
	{     
		Write-Entry "Object is not unique in Active Directory. Found $($DN.count) matches." $global:LogTypeError
		return $null
	}
	else
	{
		Write-Entry "Object found in Active Directory, $($DN[0].path)." $global:LogTypeInfo
		return $DN
	}

}

function Move-ADObject {
    <#
        .DESCRIPTION
            Move an Active Directory object to a different OU
            
        .PARAMETER DN
            System.DirectoryServices object to move
		
        .PARAMETER NewOU
            Destination OU for object
		
		.OUTPUT/RETURN
			Nothing
		
        .EXAMPLE
            Move-ADObject $currentComputer.path "OU=Computers,OU=Brisbane,DC=HQ,DC=LOCAL"
		
		.NOTES
			Author:		David Bubb
			Created:	2010/13/12
	#>

    [CmdletBinding()]
    PARAM
    (
        [Parameter(Position=1, Mandatory=$true)] [string]$DN,
		[Parameter(Position=2, Mandatory=$true)] [string]$NewOU
    )
	
	try {
		$adObject = [adsi]("$DN")
		$objectNewOU = [adsi]("LDAP://$NewOU")
		$adObject.PSBase.MoveTo($objectNewOU)
		Write-Entry "$DN moved to $NewOU" $global:LogTypeInfo
		
	} catch {
		Write-Entry "Error moving object $_" $global:LogTypeError
	}

}

function Get-Gateway {
    <#
        .DESCRIPTION
            Get the IPv4 default gateway of the active (IP enabled) network adapter
        
		.OUTPUT/RETURN
			Return = IP Gateway string
		
        .EXAMPLE
            Get-Gateway
		
		.NOTES
			Author:		David Bubb
			Created:	2010/14/12
	#>

	$adapters = get-wmiobject -query ("select IPAddress,DefaultIPGateway from Win32_NetworkAdapterConfiguration where IPEnabled=true")
	foreach ($adapter in $adapters) {
		if (($adapter.IPAddress -ne $NULL) -and ($adapter.DefaultIPGateway -ne $NULL)) {
			return $adapter.DefaultIPGateway
		}
	}
}

function Get-Laptop {
    <#
        .DESCRIPTION
            Determine if machine is a laptop/portable device
            
        .PARAMETER Computer
            Computer to test (default is local host)
		
		.OUTPUT/RETURN
			Return = True|False
		
        .EXAMPLE
            Get-Laptop
		
		.NOTES
			Author:		David Bubb
			Created:	2010/14/12
	#>

	[CmdletBinding()]
    PARAM
    (
		[Parameter(Position=1)] [string]$computer = "localhost"
	)
	
	$isLaptop = $false
	if(Get-WmiObject -Class win32_systemenclosure -ComputerName $computer |	Where-Object { $_.chassistypes -eq 8 -or $_.chassistypes -eq 9 -or $_.chassistypes -eq 10 -or $_.chassistypes -eq 12 -or $_.chassistypes -eq 14 -or $_.chassistypes -eq 18 -or $_.chassistypes -eq 21}) {
		$isLaptop = $true
	}
	if(Get-WmiObject -Class win32_battery -ComputerName $computer) {
		$isLaptop = $true
	}
	return $isLaptop
}

function Create-Directory {
    <#
        .DESCRIPTION
            Creates a specified directory and varifies that it then exists
            
        .PARAMETER Path
            The Path/Directory to create
		
		.OUTPUT/RETURN
			Nothing
		
        .EXAMPLE
            Create-Directory "C:\Support"
		
		.NOTES
			Author:		Jacob Hodges
			Created:	2010/10/07
	#>

	[CmdletBinding()]
    PARAM
    (
        [Parameter(Position=1, Mandatory=$true)] [string]$path
    )
	
	Write-Entry "Attempting to create $path" $global:LogTypeInfo
	New-Item -ItemType Directory -Path $path -ErrorAction SilentlyContinue
		
	# Make sure it exist, if not write an error
	if(Test-Path $path -PathType Container)	{
		Write-Entry "Successfully created $path." $global:LogTypeInfo
	} else {
		Write-Entry "Error creating $path." $global:LogTypeError
	}
}

function Rename-Directory {
    <#
        .DESCRIPTION
            Renames a specified directory and varifies that it then exists
            
        .PARAMETER Path
            The Path/Directory to rename
			The new name
		
		.OUTPUT/RETURN
			Nothing
		
        .EXAMPLE
            Rename-Directory "C:\Support" "NewSupport"
		
		.NOTES
			Author:		David Bubb
			Created:	2012/11/26
	#>

	[CmdletBinding()]
    PARAM
    (
        [Parameter(Position=1, Mandatory=$true)] [string]$path,
		[Parameter(Position=2, Mandatory=$true)] [string]$newname
    )
	
	Write-Entry "Attempting to rename $path to $newname" $global:LogTypeInfo
	Rename-Item -Path $path -NewName $newname -ErrorAction SilentlyContinue
		
	# Make sure it exist, if not write an error
	$parent = Split-Path $path -Parent
	if(Test-Path "$parent\$newname" -PathType Container)	{
		Write-Entry "Successfully renamed $path to $newname." $global:LogTypeInfo
	} else {
		Write-Entry "Error renaming $path." $global:LogTypeError
	}
}

function Move-Directory {
    <#
        .DESCRIPTION
            Moves a specified directories content to a different folder.
			Creates new location if it does not exist already.
            
        .PARAMETER Path
            The Path/Directory to move
			Target Path/Directory
		
		.OUTPUT/RETURN
			Nothing
		
        .EXAMPLE
            Move-Directory "C:\Support" "C:\NewSupport"
		
		.NOTES
			Author:		David Bubb
			Created:	2012/11/26
	#>

	[CmdletBinding()]
    PARAM
    (
        [Parameter(Position=1, Mandatory=$true)] [string]$sourcepath,
		[Parameter(Position=2, Mandatory=$true)] [string]$destinationpath
    )
	
	If (Test-Path $sourcepath -PathType Container) {
		If (-not(Test-Path $destinationpath -PathType Container)) {
			Create-Directory $destinationpath
		}
		Write-Entry "Attempting to move $sourcepath to $destinationpath" $global:LogTypeInfo
		Move-Item -Path $path -Destination $destinationpath -ErrorAction SilentlyContinue
			
		# Make sure destination contains some files/folders, if not write an error
		if((Get-ChildItem $destinationpath).Count -gt 0)	{
			Write-Entry "Successfully moved $sourcepath to $destinationpath." $global:LogTypeInfo
		} else {
			Write-Entry "Error moving $sourcepath." $global:LogTypeError
		}
	} else {
		Write-Entry "Error moving $sourcepath, it does not exist." $global:LogTypeError
	}
}

function Configure-ACL {
    <#
        .DESCRIPTION
            Set NTFS ACL on a file or folder
            
        .PARAMETER path
            Path to file or folder
		
        .PARAMETER aclmod
            sddl format ACL to apply to $path
		
		.OUTPUT/RETURN
			Nothing
		
        .EXAMPLE
            Configure-ACL "C:\Support" "O:BAG:S-1-5-21-3169394314-1078846113-2570888003-513D:PAI(A;OICI;FA;;;SY)(A;OICI;FA;;;BA)(A;OICI;0x1200a9;;;BU)"
		
		.NOTES
			Author:		David Bubb
			Created:	2011/01/19
	#>
	<#
	sddl format ACL
	#>
	[CmdletBinding()]
    PARAM
    (
        [Parameter(Position=1, Mandatory=$true)] [string]$path,
		[Parameter(Position=2, Mandatory=$true)] [string]$aclmod
    )
	
	Write-Entry "Attempting to configure ACL on $path" $global:LogTypeInfo

	# Set ACL on $path
	$acl = (get-acl $path)
	$acl.SetSecurityDescriptorSddlForm($aclmod) 
	set-acl $path $acl
	If (-not $?) {
		# ACL failed to apply
		Write-Entry "Error setting ACL $($aclmod) on $($path)." $global:LogTypeError
	} Else {
		Write-Entry "ACL configured for $($path)" $global:LogTypeInfo
	}
}

function Create-Shortcut {
    <#
        .DESCRIPTION
            Creates a shortcut based on supplied information
            
        .PARAMETER p
            Path to place shortcut
		
        .PARAMETER n
            Name for shortcut
		
		.PARAMETER t
			Target of shortcut
		
		.PARAMETER ta
			Target arguments (if applicable)
		
		.PARAMETER ic
			Path to icon file
		
		.PARAMETER d
			Description
		
		.PARAMETER w
			Windows type (max/min, default = normal)
		
		.OUTPUT/RETURN
			Nothing
		
        .EXAMPLE
            Create-Shortcut "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup" "Message" "C:\Windows\Notepad.exe" "\\domain.local\netlogon\dailymsg.txt" "C:\Windows\System32\shell32.dll, 23" "Open the daily message" "max"
		
		.NOTES
			Author:		David Bubb
			Created:	2011/01/17
	#>

	[CmdletBinding()]
    PARAM
    (
        [Parameter(Position=1, Mandatory=$true)] [string]$p,
		[Parameter(Position=2, Mandatory=$true)] [string]$n,
		[Parameter(Position=3, Mandatory=$true)] [string]$t,
		[Parameter(Position=4)] [string]$ta = $null,
		[Parameter(Position=5, Mandatory=$true)] [string]$ic = $null,
		[Parameter(Position=6)] [string]$d = $null,
		[Parameter(Position=7)] [string]$w = $null
    )
	
	Write-Entry "Attempting to create shortcut $($n) to $($t) in $($p)" $global:LogTypeInfo
	Write-Entry "Full paramater list; path=$($p) name=$($n) target=$($t) arguments=$($ta) icon=$($ic) description=$($d) window-style=$($w)" $global:LogTypeInfo

	# // Check to see if the shortcut already exists
	If (Test-Path "$($p)\$($n).lnk" -pathtype leaf) {
		Write-Entry "Shortcut $($p)\$($n).lnk already exists, exiting" $global:LogTypeWarning
		exit
	}
	
	If (Test-Path $p) {
		# // Sort out the window style
		$iNormalWindow = 1
		$iMaxWindow = 3
		$iMinWindow = 7
		If ($w -ne $null) {
			switch ($w.string.tolower) {
				max {$windowstyle = $iMaxWindow}
				min {$windowstyle = $iMinWindow}
			}
		} Else {
			$windowstyle = $iNormalWindow
		}
		
		$wshshell = New-Object -ComObject WScript.Shell
		$lnk = $wshshell.CreateShortcut("$($p)\$($n).lnk")
		If ($ic -ne $null) {$lnk.IconLocation = $ic}
		$lnk.Description = "$($d)"
		$lnk.WindowStyle = $windowstyle
		$lnk.TargetPath = "$($t)"
		If (-not $?) {
			# // Shortcut was not created successfully
			Write-Entry "Shortcut creation failed, invalid target" $global:LogTypeError
		} Else {
			$lnk.Arguments = "$($ta)"
			$lnk.Save()
			If (-not $?) {
				Write-Entry "Shortcut creation failed" $global:LogTypeError
			} Else {
				Write-Entry "Shortcut creation successful" $global:LogTypeInfo
			}
		}
		
		# // Clean up variables
		Clear-Variable iNormalWindow
		Clear-Variable iMaxWindow
		Clear-Variable iMinWindow
		Clear-Variable wshshell
		Clear-Variable lnk
	} Else {
		Write-Entry "Path for shortcut does not exist, unable to create" $global:LogTypeError
	}
	
	# // Clean up variables
	Clear-Variable p
	Clear-Variable n
	Clear-Variable t
	Clear-Variable ta
	Clear-Variable ic
	Clear-Variable d
	Clear-Variable w
}

# This function is superceded by Change-Service-Startup
function Disable-Service {
    <#
        .DESCRIPTION
            Disables a Windows Service.
            
        .PARAMETER Service
            Name of service to be disabled
		
		.OUTPUT/RETURN
			Nothing
		
        .EXAMPLE
            Disable-Service "WinDefend"
		
		.NOTES
			Author:		David Bubb
			Created:	2013/11/29
	#>

	[CmdletBinding()]
    PARAM
    (
        [Parameter(Position=1, Mandatory=$true)] [string]$service
    )

	Get-Service $service
	# Test to see if the action succeeded
	If ($?) {
		Set-Service $service -startupType Disabled
		Write-Entry "Successfully disabled $service." $global:LogTypeInfo
	} Else {
		Write-Entry "$service service not found." $global:LogTypeError
	}
}

function Write-Registry {
    <#
        .DESCRIPTION
            Set a value in the registry.
            
        .PARAMETER RegKey
            Root key of registry
			HKLM or HKCU or HKU
			
		.PARAMETER RegPath
			Registry path
			
		.PARAMETER RegValue
			Registry value
		
		.PARAMETER RegData
			Data to place in registry value
			Specify DWord in decimal format
		
		.PARAMETER RegType
			Type of registry value
			String or ExpandString or MultiString or Binary or DWord or QWord
		
		.OUTPUT/RETURN
			Nothing
		
        .EXAMPLE
            Write-Registry "HKLM" "SOFTWARE\Microsoft" "MyValue" "0" "DWord"
		
		.NOTES
			Author:		David Bubb
			Created:	2013/11/29
	#>

	[CmdletBinding()]
    PARAM
    (
        [Parameter(Position=1, Mandatory=$true)] [string]$RegKey,
		[Parameter(Position=2, Mandatory=$true)] [string]$RegPath,
		[Parameter(Position=3, Mandatory=$true)] [string]$RegValue,
		[Parameter(Position=4, Mandatory=$true)] [string]$RegData,
		[Parameter(Position=5, Mandatory=$true)] [string]$RegType
    )
	
	# Create a HKU drive in case we need it
	New-PSDrive -Name HKU -PSProvider Registry -Root HKEY_USERS

	# Create the path, in case it doesn't already exist
	$PathBits = $RegPath.Split("\")
	Foreach ($Bit in $PathBits) {
		$BuildPath += "\$Bit"
		New-Item "$($RegKey):\$($BuildPath)" -ErrorAction SilentlyContinue
	}	
	Set-ItemProperty "$($RegKey):\$($RegPath)" -name $RegValue -type $RegType -value $RegData
	# Test to see if the Registry action succeeded
	If ($?) {
		Write-Entry "Successfully set $RegKey\$RegPath, $RegValue = $RegData (type $RegType)." $global:LogTypeInfo
	} Else {
		Write-Entry "Failed to set $RegKey\$RegPath, $RegValue = $RegData (type $RegType)." $global:LogTypeError
	}
}

function Disable-SchedTask {
    <#
        .DESCRIPTION
            Disables a Windows Scheduled Task.
            
        .PARAMETER Task
            Name of task to be disabled
		
		.OUTPUT/RETURN
			Nothing
		
        .EXAMPLE
            Disable-SchedTask "\Microsoft\Windows\Defrag\ScheduledDefrag"
		
		.NOTES
			Author:		David Bubb
			Created:	2013/11/29
	#>

	[CmdletBinding()]
    PARAM
    (
        [Parameter(Position=1, Mandatory=$true)] [string]$task
    )

	If ([System.Environment]::OSVersion.Version.Major -eq 6 -and [System.Environment]::OSVersion.Version.Minor -gt 1) {
		Disable-ScheduledTask -TaskName $task
			# Test to see if the Registry action succeeded
		If ($?) {
			Write-Entry "Successfully disabled task $task." $global:LogTypeInfo
		} Else {
			Write-Entry "Failed to disabled task $task." $global:LogTypeError
		}
	} Else {
		schtasks /change /tn "$task" /disable
		If ($?) {
			Write-Entry "Successfully disabled task $task." $global:LogTypeInfo
		} Else {
			Write-Entry "Failed to disabled task $task." $global:LogTypeError
		}
	}
}

function Halt-Service {
    <#
        .DESCRIPTION
            Stops a Windows Service.
            
        .PARAMETER Service
            Name of service to be stopped
		
		.OUTPUT/RETURN
			Nothing
		
        .EXAMPLE
            Halt-Service "WinDefend"
		
		.NOTES
			Author:		David Bubb
			Created:	2013/12/11
	#>

	[CmdletBinding()]
    PARAM
    (
        [Parameter(Position=1, Mandatory=$true)] [string]$service
    )

	Get-Service $service
	# Test to see if the action succeeded
	If ($?) {
		Stop-Service $service -Force
		If ($?) {
			Write-Entry "Successfully stopped $service." $global:LogTypeInfo
		} Else {
			Write-Entry "$service could not be stopped." $global:LogTypeError
		}
	} Else {
		Write-Entry "$service service not found." $global:LogTypeError
	}
}

function Delete-RegistryValue {
    <#
        .DESCRIPTION
            Delete a value from the registry.
            
        .PARAMETER RegKey
            Root key of registry
			HKLM or HKCU or HKU
			
		.PARAMETER RegPath
			Registry path
			
		.PARAMETER RegValue
			Registry value
				
		.OUTPUT/RETURN
			Nothing
		
        .EXAMPLE
            Delete-Registry "HKLM" "SOFTWARE\Microsoft" "MyValue"
		
		.NOTES
			Author:		David Bubb
			Created:	2013/12/11
	#>

	[CmdletBinding()]
    PARAM
    (
        [Parameter(Position=1, Mandatory=$true)] [string]$RegKey,
		[Parameter(Position=2, Mandatory=$true)] [string]$RegPath,
		[Parameter(Position=3, Mandatory=$true)] [string]$RegValue
    )
	
	# Create a HKU drive in case we need it
	New-PSDrive -Name HKU -PSProvider Registry -Root HKEY_USERS

	Remove-ItemProperty "$($RegKey):\$($RegPath)" -name $RegValue -Force
	# Test to see if the Registry action succeeded
	If ($?) {
		Write-Entry "Successfully removed $RegKey\$RegPath, $RegValue." $global:LogTypeInfo
	} Else {
		Write-Entry "Failed to remove $RegKey\$RegPath, $RegValue." $global:LogTypeError
	}
}

function Delete-Files {
    <#
        .DESCRIPTION
            Delete files from a specified folder based on a specified file pattern.
            
        .PARAMETER TargetFolder
            Folder to find files in
			
		.PARAMETER FilePattern
			What file or files to delete based on file name or pattern
			
		.PARAMETER Recurse
			Enable deletion of files in folder and subfolders
				
		.OUTPUT/RETURN
			Nothing
		
        .EXAMPLE
            Delete-Files -TargetFolder C:\Logs -FilePattern *.log
			Delete-Files -TargetFolder C:\Backup -FilePattern *.* -Recurse
		
		.NOTES
			Author:		David Bubb
			Created:	2013/12/11
	#>

	[CmdletBinding()]
    PARAM
    (
        [Parameter(Position=1, Mandatory=$true)] [string]$TargetFolder,
		[Parameter(Position=2, Mandatory=$true)] [string]$FilePattern,
		[switch]$Recurse
    )
	
	# Get files based on lastwrite filter in specified folder
	If ($Recurse) {
		$Files = Get-Childitem "$($TargetFolder)\*" -Include $FilePattern -Recurse
	} Else {
		$Files = Get-Childitem "$($TargetFolder)\*" -Include $FilePattern
	}

	Foreach ($File in $Files) {
		If ($File -ne $NULL) {
			Remove-Item $File.FullName -Force
			If ($?) {
				Write-Entry "Successfully deleted $($File.FullName)." $global:LogTypeInfo
			} Else {
				Write-Entry "Failed to deleted $($File.FullName)." $global:LogTypeError
			}
		}
		Else {
			Write-Entry "No more files to delete." $global:LogTypeInfo
		}
	}
}

function Change-Service-Startup {
    <#
        .DESCRIPTION
            Change the startup type for a Windows service.
            
        .PARAMETER Service
            Name of service to be disabled
		
		.OUTPUT/RETURN
			Nothing
		
        .EXAMPLE
            Change-Service-Startup "WinDefend" "disabled"
			Change-Service-Startup "SmcService" "manual"
			Change-Service-Startup "WinRM" "automatic"
		
		.NOTES
			Author:		David Bubb
			Created:	2013/12/11
	#>

	[CmdletBinding()]
    PARAM
    (
        [Parameter(Position=1, Mandatory=$true)] [string]$service,
		[Parameter(Position=2, Mandatory=$true)] [string]$starttype
    )

	Get-Service $service
	# Test to see if the action succeeded
	If ($?) {
		Set-Service $service -startupType $starttype
		Write-Entry "Successfully set startup type of $service service to $starttype." $global:LogTypeInfo
	} Else {
		Write-Entry "$service service not found." $global:LogTypeError
	}
}

Function Show-MessageBox{ 
	<# 
		.SYNOPSIS  
		  Displays a MessageBox using Windows WinForms 
		   
		.Description 
			  This function helps display a custom Message box with the options to set 
			  what Icons and buttons to use. By Default without using any of the optional 
			  parameters you will get a generic message box with the OK button. 
		   
		.Parameter Msg 
			Mandatory: This item is the message that will be displayed in the body 
			of the message box form. 
			Alias: M 
	 
		.Parameter Title 
			Optional: This item is the message that will be displayed in the title 
			field. By default this field is blank unless other text is specified. 
			Alias: T 
	 
		.Parameter OkCancel 
			Optional:This switch will display the Ok and Cancel buttons. 
			Alias: OC 
	 
		.Parameter AbortRetryIgnore 
			Optional:This switch will display the Abort Retry and Ignore buttons. 
			Alias: ARI 
	 
		.Parameter YesNoCancel 
			Optional: This switch will display the Yes No and Cancel buttons. 
			Alias: YNC 
	 
		.Parameter YesNo 
			Optional: This switch will display the Yes and No buttons. 
			Alias: YN 
	 
		.Parameter RetryCancel 
			Optional: This switch will display the Retry and Cancel buttons. 
			Alias: RC 
	 
		.Parameter Critical 
			Optional: This switch will display Windows Critical Icon. 
			Alias: C 
	 
		.Parameter Question 
			Optional: This switch will display Windows Question Icon. 
			Alias: Q 
	 
		.Parameter Warning 
			Optional: This switch will display Windows Warning Icon. 
			Alias: W 
	 
		.Parameter Informational 
			Optional: This switch will display Windows Informational Icon. 
			Alias: I 
	 
		.Example 
			Show-MessageBox -Msg "This is the default message box" 
			 
			This example creates a generic message box with no title and just the  
			OK button. 
		 
		.Example 
			$A = Show-MessageBox -Msg "This is the default message box" -YN -Q 
			 
			if ($A -eq "OK" )  
			{ 
				..do something  
			}  
			else  
			{  
			 ..do something else  
			}  
	 
			This example creates a msgbox with the Yes and No button and the 
			Question Icon. Once the message box is displayed it creates the A varible 
			with the message box selection choosen.Once the message box is done you  
			can use an if statement to finish the script. 
			 
		.Notes 
			Version: 1.0 
			Created By: Zachary Shupp 
			Email: Zach.Shupp@outlook.com 
			Date: 9/23/2013 
			Purpose/Change:    Initial function development 
	 
			Version 1.1 
			Created By Zachary Shupp 
			Email: Zach.Shupp@outlook.com 
			Date: 12/13/2013 
			Purpose/Change: Added Switches for the form Type and Icon to make it easier to use. 
			 
		.Link 
			http://msdn.microsoft.com/en-us/library/system.windows.forms.messagebox.aspx 
			 
	#>
    Param( 
    [Parameter(Mandatory=$True)][Alias('M')][String]$Msg, 
    [Parameter(Mandatory=$False)][Alias('T')][String]$Title = "", 
    [Parameter(Mandatory=$False)][Alias('OC')][Switch]$OkCancel, 
    [Parameter(Mandatory=$False)][Alias('OCI')][Switch]$AbortRetryIgnore, 
    [Parameter(Mandatory=$False)][Alias('YNC')][Switch]$YesNoCancel, 
    [Parameter(Mandatory=$False)][Alias('YN')][Switch]$YesNo, 
    [Parameter(Mandatory=$False)][Alias('RC')][Switch]$RetryCancel, 
    [Parameter(Mandatory=$False)][Alias('C')][Switch]$Critical, 
    [Parameter(Mandatory=$False)][Alias('Q')][Switch]$Question, 
    [Parameter(Mandatory=$False)][Alias('W')][Switch]$Warning, 
    [Parameter(Mandatory=$False)][Alias('I')][Switch]$Informational) 
 
    #Set Message Box Style 
    IF($OkCancel){$Type = 1} 
    Elseif($AbortRetryIgnore){$Type = 2} 
    Elseif($YesNoCancel){$Type = 3} 
    Elseif($YesNo){$Type = 4} 
    Elseif($RetryCancel){$Type = 5} 
    Else{$Type = 0} 
     
    #Set Message box Icon 
    If($Critical){$Icon = 16} 
    ElseIf($Question){$Icon = 32} 
    Elseif($Warning){$Icon = 48} 
    Elseif($Informational){$Icon = 64} 
    Else{$Icon = 0} 
     
    #Loads the WinForm Assembly, Out-Null hides the message while loading. 
    [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null 
 
    #Display the message with input 
    $Answer = [System.Windows.Forms.MessageBox]::Show($MSG , $TITLE, $Type, $Icon) 
     
    #Return Answer 
    Return $Answer 
}

function Get-FreeSpace {
<#
        .DESCRIPTION
            Retreive free disk space in Megabytes of the Volume containing Windows path
		
        .OUTPUT/RETURN
            Free Disk space in Megabytes
			Output = Debuggin message
		
		.EXAMPLE
            Get-FreeSpace
		
		.NOTES
			Author:		Jesse Harris
			Created:	2015/08/06
#>
[CmdLetBinding()]
Param($Volume=($env:windir).Remove(2))
    $ByteSpace = (Get-WMIObject -Query "Select FreeSpace from Win32_LogicalDisk Where DeviceID = ""$Volume""").FreeSpace
    [Int]($ByteSpace/1024/1024)
}

function Get-SCCMCacheInfo {
<#
        .DESCRIPTION
            Retreive SCCM Cache object
		
        .OUTPUT/RETURN
            SCCM Cache Object
			Output = Debuggin message
		
		.EXAMPLE
            Get-SCCMCacheInfo
		
		.NOTES
			Author:		Jesse Harris
			Created:	2015/08/06
	#>
	
    [CmdletBinding()]
    PARAM()

    #Ensure Service is still active
    Write-Entry "Testing Config Manager Client service is active" $global:LogTypeInfo
    $Service = Get-Service -Name CCMExec
    # Test to see if the action succeeded
	If ($Service.Status -eq 'Running') {
        $OUIResource = New-Object -ComObject UIResource.UIResourceMgr
        If ($?) {
            Write-Entry "Succesfully retreived SCCM Cache object" $global:LogTypeInfo
            $OUIResource.GetCacheInfo()
        } Else {
            Write-Entry "Unable to retrieve SCCM Cache object" $global:LogTypeError
        }      
    } Else {
        Write-Entry "CCMExec service is not currently present or running" $global:LogTypeError
    }
}

function Clear-SCCMCache {
    <#
        .DESCRIPTION
            Gracefully clear SCCM Cache while client is still active
		
        .OUTPUT/RETURN
			Output = Debuggin message
		
		.EXAMPLE
            Clear-SCCMCache
		
		.NOTES
			Author:		Jesse Harris
			Created:	2015/08/06
	#>
	
    [CmdletBinding()]
    PARAM()

    #Ensure Service is still active
    $CacheInfo = Get-SCCMCacheInfo
    # Test to see if the action succeeded
	If ($CacheInfo) {
        $BeforeSpace = Get-FreeSpace
        ForEach ($CacheObj in $CacheInfo.GetCacheElements()) {
            If ($CacheObj.ReferenceCount -eq 0) {
                Try{
                    $CacheInfo.DeleteCacheElement($CacheObj.CacheElementId)
                } Catch {
                    Write-Entry "Unable to remove item. Possibly in use." $global:LogTypeError
                }
            }
        }
        $AfterSpace = Get-FreeSpace
        $SavedSpace = $AfterSpace - $BeforeSpace
        Write-Entry "Sucessfuly saved $SavedSpace Megabytes" $global:LogTypeInfo
    } Else {
        Write-Entry "Unable to clear SCCM Cache" $global:LogTypeError
    }
}

function Cache-AppVPackages {
    <#
        .DESCRIPTION
            Mount all appv packages to the local hard disk
		
        .OUTPUT/RETURN
			Output = Debuggin message
		
		.EXAMPLE
            Cache-AppVPackages
		
		.NOTES
			Author:		Jesse Harris
			Created:	2015/08/28
	#>
	
    [CmdletBinding()]
    PARAM()

    Import-Module 'C:\Program Files\Microsoft Application Virtualization\Client\AppvClient\AppvClient.psd1'

    #Ensure Service is still active
    $APPVService = Get-Service -Name AppVClient
    # Test to see if the action succeeded
	If ($APPVService) {
        Write-Entry "AppV Service is live" $global:LogTypeInfo
        $Packages = Get-AppvClientPackage -All
        ForEach ($Package in $Packages) {
            If ($Package.PercentLoaded -ne 100) {
                Write-Entry "$($Package.Name) is not fully loaded. Attempting mount." $global:LogTypeInfo
                Try {
                    Mount-AppvClientPackage -PackageId $Package.PackageId -VersionId $Package.VersionId
                    Write-Entry "Succesfully mounted $($Package.Name)" $global:LogTypeInfo
                } Catch {
                    Write-Entry "Unable to mount $($Package.Name)" $global:LogTypeError
                }
            } Else {
                Write-Entry "$($Package.Name) is already loaded." $global:LogTypeInfo
            }
        }
    } Else {
        Write-Entry "AppV Service not available." $global:LogTypeError
    }
}

function Remove-UnusedAppVPackages {
    <#
        .DESCRIPTION
            Clear unpiblished packages
		
        .OUTPUT/RETURN
			Output = Debuggin message
		
		.EXAMPLE
            Remove-UnusedAppVPackages
		
		.NOTES
			Author:		Jesse Harris
			Created:	2015/08/28
	#>
	
    [CmdletBinding()]
    PARAM()

    Import-Module 'C:\Program Files\Microsoft Application Virtualization\Client\AppvClient\AppvClient.psd1'

    #Ensure Service is still active
    $APPVService = Get-Service -Name AppVClient
    # Test to see if the action succeeded
	If ($APPVService) {
        Write-Entry "AppV Service is live" $global:LogTypeInfo
        Import-Module 'C:\Program Files\Microsoft Application Virtualization\Client\AppvClient\AppvClient.psd1'
        $Packages = Get-AppvClientPackage -All
        ForEach ($Package in $Packages) {
            If (-Not $Package.IsPublishedGlobally) {
                Write-Entry "$($Package.Name) is not published globally. Attempting removal." $global:LogTypeInfo
                Try {
                    Remove-AppvClientPackage -PackageId $Package.PackageId -VersionId $Package.VersionId
                    Write-Entry "Succesfully removed $($Package.Name)" $global:LogTypeInfo
                } Catch {
                    Write-Entry "Unable to remove $($Package.Name)" $global:LogTypeError
                }
            } Else {
                Write-Entry "$($Package.Name) is published globally. Skipping." $global:LogTypeInfo
            }
        }
    } Else {
        Write-Entry "AppV Service not available." $global:LogTypeError
    }
}

function Update-PVDInventory {
    <#
        .DESCRIPTION
            Update PVD Inventory
		
        .OUTPUT/RETURN
			Output = Debuggin message
		
		.EXAMPLE
            Update-PVDInventory
		
		.NOTES
			Author:		Jesse Harris
			Created:	2015/08/28
	#>
	
    [CmdletBinding()]
    PARAM()

    $PVDExecutable = Get-Item -Path 'C:\Program Files\Citrix\personal vDisk\bin\CtxPvD.exe'
    $PVDRegistrySetting = Get-ItemProperty -Path 'HKLM:\Software\Citrix\personal vDisk\Config' -Name InterceptShutdown
    If ($PVDRegistrySetting.InterceptShutdown -eq 0) {
        #PVD Inventory hasn't been run. Lets run it now
        Write-Entry "InterceptShutdown is set to 0, meaning we haven't run on shutdown" $global:LogTypeInfo
        $PVDResult = Start-Process -FilePath $PVDExecutable -ArgumentList "-s inventoryonly" -Wait -PassThru
        If ($PVDResult.ExitCode -ne 0) {
            Write-Entry "PvD Inventory closed with exit code $($PVDResult.ExitCode)" $global:LogTypeError
        } Else {
            Write-Entry "PvD Inventory closed with exit code $($PVDResult.ExitCode)" $global:LogTypeInfo
        }
    } Else {
        Write-Entry "InterceptShutdown is set to 1. It might have already been run." $global:LogTypeInfo
    }
}

function Disk-Cleanup {
    <#
        .DESCRIPTION
            Run Disk Cleanup automated
		
        .OUTPUT/RETURN
			Output = Debuggin message
		
		.EXAMPLE
            Disk-Cleanup
		
		.NOTES
			Author:		Jesse Harris
			Created:	2015/08/28
	#>
	
    [CmdletBinding()]
    PARAM()

    $BeforeSpace = Get-FreeSpace
    Write-Entry "Running disk cleanup" $global:LogTypeInfo
    Start-Process -FilePath C:\Windows\system32\cleanmgr.exe -ArgumentList "/sagerun:1" -Wait
    $AfterSpace = Get-FreeSpace
    $SavedSpace = $AfterSpace - $BeforeSpace
    Write-Entry "Sucessfuly saved $SavedSpace Megabytes" $global:LogTypeInfo
}

function Remove-APPVLocalVFSSecured {
    <#
        .DESCRIPTION
            Clear the localvfssecured registry key 
		
        .OUTPUT/RETURN
			Output = Debuggin message
		
		.EXAMPLE
            Remove-APPVLocalVFSSecured
		
		.NOTES
			Author:		Jesse Harris
			Created:	2015/08/31
	#>
	
    [CmdletBinding()]
    PARAM()

    #Check for the existence of the key
    $LOCALVFSSecuredPath = 'HKLM:\Software\Microsoft\AppV\Client\Virtualization\LocalVFSSecuredUsers'
    $LOCALVFSSecured = Get-Item -Path $LOCALVFSSecuredPath
    If ($LOCALVFSSecured) {
        ForEach ($ItemProperty in $LOCALVFSSecured.Property) {
            If ($ItemProperty -ne 'S-1-5-18') {
                Write-Entry "Removing SID $ItemProperty" $global:LogTypeInfo        
                Remove-ItemProperty -Path $LOCALVFSSecuredPath -Name $ItemProperty
                If ($? -eq $False) {
                    Write-Entry "Unable to remove SID: $ItemProperty" $global:LogTypeError
                }
            }
        }
    }
}