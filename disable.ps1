#####################################################
# HelloID-Conn-Target-SkypeFB-OnPremises-Disable
#
# Version: 1.0.0
#####################################################
#Region functions

function New-SkypeFBSession {
    <#
    .SYNOPSIS
    Creates a new Skype for business remote session

    .DESCRIPTION
    Creates a new Skype for business remote PowerShell session

    .PARAMETER ComputerName
    The hostname or IP address of the server that hosts the Skype for business environment

    .PARAMETER UserName
    The username to connect to the server. Make sure the user has administrative rights

    .PARAMETER Password
    The password

    .EXAMPLE
    #>
    [CmdletBinding()]
    param (
        [String]
        $ComputerName,

        [String]
        $UserName,

        [String]
        $Password
    )

    try {
        $securePassword = $Password | ConvertTo-SecureString -AsPlainText -Force
        [pscredential]$credentials = New-Object System.Management.Automation.PSCredential ($UserName, $securePassword)
        New-PSSession -ComputerName $computerName -Credential $credentials
    } catch {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}

function Invoke-SkypeFBCommand {
    <#
    .SYNOPSIS
    Invokes a Sfb command within the remote session

    .DESCRIPTION
    Invokes a Sfb command within the remote session. This function is used to execute a script containing the logic to perform actions
    on the SkypeForBusiness server. The script is send as a scriptblock

    .PARAMETER Session
    The PowerShell session created at the 'New-SkypeFBSession'

    .PARAMETER Command
    The command(s) or scriptblock to execute within the remote session

    .EXAMPLE
    #>
    [CmdletBinding()]
    param (
        [System.Management.Automation.Runspaces.PSSession]
        $Session,

        [ScriptBlock]
        $Command
    )

    try {
        Invoke-Command -Session $Session -ScriptBlock $Command
    } catch {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}

function Remove-ResponseGroupMember {
    <#
    .SYNOPSIS
    Removes single or multiple users from specific or all response groups.

    .Description
    Remove-RgsMember uses takes users via the pipeline and removes them from the desired response groups. When specifying the groups, it will remove all users from just those groups. When specifying all groups with the switch, it wil lgo through all groups they are member of and remove them.

    .PARAMETER User
    The user(s) to remove. The user needs to be a SIP address in the form "sip:user@sipdomain". This is required and can utilize the pipeline.

    .PARAMETER ResponseGroup
    Used to specify individual response groups. Multiple response groups can be specified and should be specified by their 'Name' property. This parameter cannot be used with the AllGroups switch.

    .PARAMETER AllGroups
    Used to remove the specified users from all response groups they are a member of. This parameter cannot be used with the ResponseGroup parameter.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [String[]]
        $User,

        [Parameter(ParameterSetName = "specifiedGroups")]
        [String[]]
        $ResponseGroup,

        [Parameter(ParameterSetName = "allGroups")]
        [Switch]
        $AllGroups
    )

    process {
        try {
            if($AllGroups) {
                foreach ($member in $User) {
                    $groupsContainingUser = Get-CsRgsAgentGroup | Where-Object {$_.AgentsByUri -contains $member}
                    Write-Verbose -Message "$member is in $($groupsContainingUser.Count) groups"

                    if($groupsContainingUser.Count -gt 0) {
                        foreach ($group in $groupsContainingUser) {
                            Write-Verbose -Message "Removing $member from $($group.Name)"
                            [void]$group.AgentsByUri.Remove($member)
                            Set-CsRgsAgentGroup -Instance $group
                        }
                    } else {
                        Write-Verbose "Not attempting removal as $member was not in any response groups"
                    }
                }
            }
            else {
                foreach ($member in $User) {
                    Write-Verbose -Message "Starting removals for $member"
                    foreach($group in $ResponseGroup) {
                        Write-Verbose -Message "Removing $member from $group"
                        $rg = Get-CsRgsAgentGroup -Name $group
                        [void]$rg.AgentsByUri.Remove($member)
                        Set-CsRgsAgentGroup -Instance $rg
                    }
                }
            }
        } catch {
            $PSCmdlet.ThrowTerminatingError($_)
        }
    }
}
#EndRegion

#Initialize default properties
$connectionSettings = ConvertFrom-Json $configuration
$personObject = ConvertFrom-Json $person

$auditMessage = "for person " + $personObject.DisplayName
$success = $False
$account_guid = New-Guid

#Change mapping here
$account = [PSCustomObject]@{
    displayName = $personObject.DisplayName
    userName = $personObject.UserName
    externalId = $account_guid
    samAccountName = $personObject.Accounts.MicrosoftActiveDirectory.sAMAaccountName
    userPrincipalName = $personObject.Accounts.MicrosoftActiveDirectory.userPrincipalName
}

if(-Not($dryRun -eq $True)) {
    Write-Verbose -Verbose "[SkypeForBusiness] Starting to disable user '$($account.userName)'"
    try {
        $session = New-SkypeFBSession -ComputerName $($connectionSettings.ComputerName) -UserName $($connectionSettings.UserName) -Password $($connectionSettings.Pasword)
        $command = {
            try {
                Import-Module 'C:\Program Files\Common Files\Skype for Business Server 2019\Modules\SkypeForBusiness' -Force
                Remove-ResponseGroupMember -UserPrincipalName $Using:account.userPrincipalName -AllGroups
                Disable-CsUser -Identity $Using:account.userPrincipalName

                Write-Verbose -Verbose "[SkypeForBusiness] Finished disabling user '$($account.userName)'"

                $success = $True
                $auditMessage = "[SkypeForBusiness] Finished disabling user '$($account.userName)'"
            } catch {
                throw $_
            }
        }
        Invoke-SkypeFBCommand -Session $session -Command $command
    } catch {
        $ex = $_
        $success = $false
        $auditMessage = "[SkypeForBusiness] Could not disable user '$($account.userName)' Error: $ex"
        Write-Error "[SkypeForBusiness] Could not disable user ''$($account.userName)' Error: $ex"
    }
}

#build up result
$result = [PSCustomObject]@{
	Success= $success
	AccountReference= $account_guid
	AuditDetails=$auditMessage
    Account = $account

    # Optionally return data for use in other systems
    ExportData = [PSCustomObject]@{
        displayName = $account.DisplayName
        userName = $account.UserName
        externalId = $account_guid
    }
}

#send result back
Write-Output $result | ConvertTo-Json -Depth 10
