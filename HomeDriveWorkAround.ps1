<#
.DESCRIPTION
    [THIS SCRIPT MUST BE RUN AS THE ENDUSER BECAUSE IT CHECKS AND MODIFIES THE USER'S REGISTRY.  IT CANNOT BE RUN AS AN ADMIN.]

    The script will update HKCU Shell Folders and User Shell Folders (specifically 'Personal','My Pictures', 'Favorites', and
    their GUID equivalents) to point to the correct network share UNC based on their HomeDrive group membership.  Users must be
    a member of one, and only one, of these HomeDrive groups.  The base UNC path for each group is listed.

        HomeDriveBK = \\BKCALL02\Profiles\
        HomeDriveBL = \\BLFILE01\
        HomeDriveCG = \\CGFILE01\
        HomeDriveNM = \\NMFILE01\
        HomeDriveSA = \\SAFILE01\
        HomeDriveSM = \\SMFILE01\
        HomeDriveVN = \\VNFILE01\

    This script is a workaround for an issue where Folder Redirect policies are not being applied, which leaves the old Shell
    Folder paths pointing to paths that include the now-missing %HomeDir% variable (due to the clearing of the HomeDirectory
    attribute in the user's AD object).

    The script will output the results to the screen as well as to $env:TEMP (the user's temp directory).  The file name will
    have a path and name format similar to:

        C:\Users\BH236250\AppData\Local\Temp\20190820-1423-ShellFoldersUpdate.CSV"   [Example]
#>

[CmdletBinding(SupportsShouldProcess=$true)]

PARAM()

#Get existing Shell Folders for current user
#User Shell Folders take precedence over Shell Folders
$ExistingShellFolders     = Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders'
$ExistingUserShellFolders = Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders'

#Get Current User's group membership
$GroupMembership = ([Security.Principal.WindowsIdentity]::GetCurrent()).Groups | ForEach-Object {
    $_.Translate([Security.Principal.NTAccount])
}

#Get HomeDrive group membership (the user should only be a member of 1 HomeDrive* group).  Warn and exit if the user is not a member of just one. (zero or 2 or more)
$HomeDriveGroups = @()
$HomeDriveGroups = $GroupMembership.Where({$_.Value -like "AERAENERGY\HomeDrive*"})
If ($HomeDriveGroups.Count -eq 0) {
    Write-Warning -Message "User is not a member of any HomeDrive group.  The local and domain groups that the user is a member of are listed below."
    $GroupMembership.Value | Sort-Object
    Throw "[ERROR] Ensure the user is a member of only one HomeDrive group for the server where the users home share is located."
} ElseIf ($HomeDriveGroups.Count -ge 2) {
    #If the user is a member of 2 or more HomeDrive* groups, break out of the script and warn the user because this is bad.
    Write-Warning -Message "User has more than one HomeDrive group membership.  User MUST ONLY have one."
    $HomeDriveGroups.Value
    Throw "[ERROR] Correct the AD user's group membership, have the user reboot the computer (or log off), and try again."
} Else {
    #If only one HomeDrive group is returned (which should be the desired state), then remove the domain prefix.
    Write-Verbose -Message "[INFO] User is a member of $HomeDriveGroups"
    $HomeDriveGroup = ($HomeDriveGroups[0] -replace "AERAENERGY\\")
}

#Group to UNC mappings (must include the trailing back slash)
$GroupToPathMappings = @{
    "HomeDriveBK" = "\\BKCALL02\Profiles\"
    "HomeDriveBL" = "\\BLFILE01\"
    "HomeDriveCG" = "\\CGFILE01\"
    "HomeDriveNM" = "\\NMFILE01\"
    "HomeDriveSA" = "\\SAFILE01\"
    "HomeDriveSM" = "\\SMFILE01\"
    "HomeDriveVN" = "\\VNFILE01\"
}

#Check if the HomeDriveGroup variable has an expected value.
If ($HomeDriveGroup -like "HomeDrive*") {
    #Get the base path by appending the username
    $BasePath = $GroupToPathMappings.$HomeDriveGroup + $env:USERNAME
} Else {
    #This shouldn't happen since an earlier check should have already returned a single HomeDrive* group.
    Throw "[ERROR] Variable HomeDriveGroup is invalid.  Its value is: `"$HomeDriveGroup`""
}


<#These are the shell folders that are in scope. Note that the GUID entries only apply to the UserShellFolders key, while the named ones apply to both.

    #Personal:
        Personal  = [home path]\My Documents
        {F42EE2D3-909F-4907-8871-4C22FC0BF756} = [home path]\My Documents

    #My Pictures:
        My Pictures = [home path]\My Documents\My Pictures
        {0DDD015D-B06C-45D5-8C4C-F59713854639}  = [home path]\My Documents\My Pictures

    #Favorites:
        Favorites = [home path]\Favorites

#>

#Form the shell folder paths
$Personal    = "$BasePath\My Documents"
$MyPictures  = "$BasePath\My Documents\My Pictures"
$Favorites   = "$BasePath\Data\Favorites"

#Set Shell Folder registry values
$OverallResults  = @()
$LogoffNecessary = $false
#Shell Folders\Personal
If ($ExistingShellFolders.Personal -ne $Personal) {
    Write-Verbose -Message "[NEEDS UPDATE] Shell Folders\Personal is currently set to " + $ExistingShellFolders.Personal + ". It should be $Personal."
    Try {
        $ExceptionMessage = $Null
        Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders' -Name "Personal" -Value $Personal -ErrorAction Stop
        $ChangeStatus = "CHANGE_SUCCESS"
        Write-Verbose -Message "[SUCCESS] Updated `'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\Personal`' to $Personal"
    } Catch {
        $ExceptionMessage = $_.Exception.Message
        Write-Warning -Message "Could not update `'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\Personal`': " + $ExceptionMessage
        $ChangeStatus = $ExceptionMessage
    } #End Try..Catch
} Else {
    $ChangeStatus = "NO_CHANGE_NECESSARY"
    Write-Verbose -Message "[SKIPPED] No change necessary for `'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\Personal`'"
} #End If..Else

$OverallResults += [pscustomobject]@{
    "ShellFolder"  = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\Personal'
    "OriginalPath" = $ExistingShellFolders.Personal
    "NewPath"      = $Personal
    "Status"       = $ChangeStatus
}


#User Shell Folders\Personal
If ($ExistingUserShellFolders.Personal -ne $Personal) {
    Write-Verbose -Message "[NEEDS UPDATE] User Shell Folders\Personal is currently set to " + $ExistingUserShellFolders.Personal + ". It should be $Personal."
    $LogoffNecessary = $true
    Try {
        $ExceptionMessage = $Null
        Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -Name "Personal" -Value $Personal -ErrorAction Stop
        $ChangeStatus = "CHANGE_SUCCESS"
        Write-Verbose -Message "[SUCCESS] Updated `'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\Personal`' to $Personal"
    } Catch {
        $ExceptionMessage = $_.Exception.Message
        Write-Warning -Message "Could not update `'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\Personal`': " + $ExceptionMessage
        $ChangeStatus = $ExceptionMessage
    } #End Try..Catch
} Else {
    $ChangeStatus = "NO_CHANGE_NECESSARY"
    Write-Verbose -Message "[SKIPPED] No change necessary for `'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\Personal`'"
} #End If..Else

$OverallResults += [pscustomobject]@{
    "ShellFolder"  = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\Personal'
    "OriginalPath" = $ExistingUserShellFolders.Personal
    "NewPath"      = $Personal
    "Status"       = $ChangeStatus
}


#User Shell Folders\{F42EE2D3-909F-4907-8871-4C22FC0BF756}
If ($ExistingUserShellFolders.'{F42EE2D3-909F-4907-8871-4C22FC0BF756}' -ne $Personal) {
    $LogoffNecessary = $true
    Write-Verbose -Message "[NEEDS UPDATE] User Shell Folders\{F42EE2D3-909F-4907-8871-4C22FC0BF756} is currently set to " + $ExistingUserShellFolders.'{F42EE2D3-909F-4907-8871-4C22FC0BF756}' + ". It should be $Personal."
    Try {
        $ExceptionMessage = $Null
        Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -Name "{F42EE2D3-909F-4907-8871-4C22FC0BF756}" -Value $Personal -ErrorAction Stop
        $ChangeStatus = "CHANGE_SUCCESS"
        Write-Verbose -Message "[SUCCESS] Updated `'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\{F42EE2D3-909F-4907-8871-4C22FC0BF756}`' to $Personal"
    } Catch {
        $ExceptionMessage = $_.Exception.Message
        Write-Warning -Message "Could not update `'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\{F42EE2D3-909F-4907-8871-4C22FC0BF756}`': " + $ExceptionMessage
        $ChangeStatus = $ExceptionMessage
    } #End Try..Catch
} Else {
    $ChangeStatus = "NO_CHANGE_NECESSARY"
    Write-Verbose -Message "[SKIPPED] No change necessary for `'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\{F42EE2D3-909F-4907-8871-4C22FC0BF756}`'"
} #End If..Else

$OverallResults += [pscustomobject]@{
    "ShellFolder"  = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\{F42EE2D3-909F-4907-8871-4C22FC0BF756}'
    "OriginalPath" = $ExistingUserShellFolders.'{F42EE2D3-909F-4907-8871-4C22FC0BF756}'
    "NewPath"      = $Personal
    "Status"       = $ChangeStatus
}


#My Pictures
#Shell Folders\My Pictures
If ($ExistingShellFolders.'My Pictures' -ne $MyPictures) {
    $LogoffNecessary = $true
    Write-Verbose -Message "[NEEDS UPDATE] Shell Folders\My Pictures is currently set to " + $ExistingShellFolders.'My Pictures' + ". It should be $MyPictures."
    Try {
        $ExceptionMessage = $Null
        Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders' -Name "My Pictures" -Value $MyPictures -ErrorAction Stop
        $ChangeStatus = "CHANGE_SUCCESS"
        Write-Verbose -Message "[SUCCESS] Updated `'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\My Pictures`' to $MyPictures"
    } Catch {
        $ExceptionMessage = $_.Exception.Message
        Write-Warning -Message "Could not update `'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\My Pictures`': " + $ExceptionMessage
        $ChangeStatus = $ExceptionMessage
    } #End Try..Catch
} Else {
    $ChangeStatus = "NO_CHANGE_NECESSARY"
    Write-Verbose -Message "[SKIPPED] No change necessary for `'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\My Pictures`'"
} #End If..Else

$OverallResults += [pscustomobject]@{
    "ShellFolder"  = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\My Pictures'
    "OriginalPath" = $ExistingShellFolders.'My Pictures'
    "NewPath"      = $MyPictures
    "Status"       = $ChangeStatus
}


#User Shell Folders\My Pictures
If ($ExistingUserShellFolders.'My Pictures' -ne $MyPictures) {
    $LogoffNecessary = $true
    Write-Verbose -Message "[NEEDS UPDATE] User Shell Folders\My Pictures is currently set to " + $ExistingUserShellFolders.'My Pictures' + ". It should be $MyPictures."
    Try {
        $ExceptionMessage = $Null
        Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -Name "My Pictures" -Value $MyPictures -ErrorAction Stop
        $ChangeStatus = "CHANGE_SUCCESS"
        Write-Verbose -Message "[SUCCESS] Updated `'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\My Pictures`' to $MyPictures"
    } Catch {
        $ExceptionMessage = $_.Exception.Message
        Write-Warning -Message "Could not update `'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\My Pictures`': " + $ExceptionMessage
        $ChangeStatus = $ExceptionMessage
    } #End Try..Catch
} Else {
    $ChangeStatus = "NO_CHANGE_NECESSARY"
    Write-Verbose -Message "[SKIPPED] No change necessary for `'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\My Pictures`'"
} #End If..Else

$OverallResults += [pscustomobject]@{
    "ShellFolder"  = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\My Pictures'
    "OriginalPath" = $ExistingUserShellFolders.'My Pictures'
    "NewPath"      = $MyPictures
    "Status"       = $ChangeStatus
}


#User Shell Folders\{0DDD015D-B06C-45D5-8C4C-F59713854639}
If ($ExistingUserShellFolders.'{0DDD015D-B06C-45D5-8C4C-F59713854639}' -ne $MyPictures) {
    $LogoffNecessary = $true
    Write-Verbose -Message "[NEEDS UPDATE] User Shell Folders\{0DDD015D-B06C-45D5-8C4C-F59713854639} is currently set to " + $ExistingUserShellFolders.'{0DDD015D-B06C-45D5-8C4C-F59713854639}' + ". It should be $MyPictures."
    Try {
        $ExceptionMessage = $Null
        Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -Name "{0DDD015D-B06C-45D5-8C4C-F59713854639}" -Value $MyPictures -ErrorAction Stop
        $ChangeStatus = "CHANGE_SUCCESS"
        Write-Verbose -Message "[SUCCESS] Updated `'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\{0DDD015D-B06C-45D5-8C4C-F59713854639}`' to $MyPictures"
    } Catch {
        $ExceptionMessage = $_.Exception.Message
        Write-Warning -Message "Could not update `'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\{0DDD015D-B06C-45D5-8C4C-F59713854639}`': " + $ExceptionMessage
        $ChangeStatus = $ExceptionMessage
    } #End Try..Catch
} Else {
    $ChangeStatus = "NO_CHANGE_NECESSARY"
    Write-Verbose -Message "[SKIPPED] No change necessary for `'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\{0DDD015D-B06C-45D5-8C4C-F59713854639}`'"
} #End If..Else

$OverallResults += [pscustomobject]@{
    "ShellFolder"  = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\{0DDD015D-B06C-45D5-8C4C-F59713854639}'
    "OriginalPath" = $ExistingUserShellFolders.'{0DDD015D-B06C-45D5-8C4C-F59713854639}'
    "NewPath"      = $MyPictures
    "Status"       = $ChangeStatus
}


#Favorites
#Shell Folders\Favorites
If ($ExistingShellFolders.Favorites -ne $Favorites) {
    $LogoffNecessary = $true
    Write-Verbose -Message "[NEEDS UPDATE] Shell Folders\Favorites is currently set to " + $ExistingShellFolders.Favorites + ". It should be $Favorites."
    Try {
        $ExceptionMessage = $Null
        Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders' -Name "Favorites" -Value $Favorites -ErrorAction Stop
        $ChangeStatus = "CHANGE_SUCCESS"
        Write-Verbose -Message "[SUCCESS] Updated `'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\Favorites`' to $Favorites"
    } Catch {
        $ExceptionMessage = $_.Exception.Message
        Write-Warning -Message "Could not update `'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\Favorites`': " + $ExceptionMessage
        $ChangeStatus = $ExceptionMessage
    } #End Try..Catch
} Else {
    $ChangeStatus = "NO_CHANGE_NECESSARY"
    Write-Verbose -Message "[SKIPPED] No change necessary for `'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\Favorites`'"
} #End If..Else

$OverallResults += [pscustomobject]@{
    "ShellFolder"  = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\Favorites'
    "OriginalPath" = $ExistingShellFolders.Favorites
    "NewPath"      = $Favorites
    "Status"       = $ChangeStatus
}


#User Shell Folders\Favorites
If ($ExistingUserShellFolders.Favorites -ne $Favorites) {
    $LogoffNecessary = $true
    Write-Verbose -Message "[NEEDS UPDATE] User Shell Folders\Favorites is currently set to " + $ExistingUserShellFolders.Favorites + ". It should be $Favorites."
    Try {
        $ExceptionMessage = $Null
        Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -Name "Favorites" -Value $Favorites -ErrorAction Stop
        $ChangeStatus = "CHANGE_SUCCESS"
        Write-Verbose -Message "[SUCCESS] Updated `'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\Favorites`' to $Favorites"
    } Catch {
        $ExceptionMessage = $_.Exception.Message
        Write-Warning -Message "Could not update `'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\Favorites`': " + $ExceptionMessage
        $ChangeStatus = $ExceptionMessage
    } #End Try..Catch
} Else {
    $ChangeStatus = "NO_CHANGE_NECESSARY"
    Write-Verbose -Message "[SKIPPED] No change necessary for `'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\Favorites`'"
} #End If..Else

$OverallResults += [pscustomobject]@{
    "ShellFolder"  = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\Favorites'
    "OriginalPath" = $ExistingUserShellFolders.Favorites
    "NewPath"      = $Favorites
    "Status"       = $ChangeStatus
}


#Show results for review
$OverallResults | Export-CSV -NoTypeInformation -Path "$env:TEMP\$(Get-Date -Format yyyyMMdd-HHmm)-ShellFoldersUpdate.csv"
Write-Host -BackgroundColor Cyan -ForegroundColor DarkBlue -Object "Results are displayed below, as well as saved to `"$env:TEMP\$(Get-Date -Format yyyyMMdd-HHmm)-ShellFoldersUpdate.csv`"."
$OverallResults | Select-object -Property Status,OriginalPath,NewPath,ShellFolder | Format-Table -AutoSize


If ($LogoffNecessary) {
    #If any changes were made, show the current values for comparison.
    $NewShellFolders     = Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders'
    $NewUserShellFolders = Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders'

    #Display the new Shell Folders and User Shell Folders values
    #This is displayed because the $OverallResults assumes that the tasks completed based on Try/Catch blocks.  This will output the current values so that they can be verified.
    Write-Host -ForegroundColor Magenta -Object "Please compare the current values below with NewPath above to verify changes were actually made."
    Write-Host -BackgroundColor Cyan -ForegroundColor DarkBlue -Object "Current Shell Folder values" -NoNewline
    $NewShellFolders | Select-Object -Property Personal,'My Pictures',Favorites | Format-List
    Write-Host -BackgroundColor Cyan -ForegroundColor DarkBlue -Object "Current User Shell Folder values" -NoNewline
    $NewUserShellFolders | Select-Object -Property Personal,'{0DDD015D-B06C-45D5-8C4C-F59713854639}','My Pictures','{F42EE2D3-909F-4907-8871-4C22FC0BF756}',Favorites | Format-List

    Write-Host -BackgroundColor Magenta -Object "If any changes were made by the script, please log off and back on for the changes to take effect."
}