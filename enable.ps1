#####################################################
# HelloID-Conn-Target-SkypeFB-OnPremises-Enable
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
    Write-Verbose -Verbose "[SkypeForBusiness] Starting to enable user '$($account.userName)'"
    try {
        $session = New-SkypeFBSession -ComputerName $($connectionSettings.ComputerName) -UserName $($connectionSettings.UserName) -Password $($connectionSettings.Pasword)
        $command = {
            try {
                Import-Module 'C:\Program Files\Common Files\Skype for Business Server 2019\Modules\SkypeForBusiness' -Force
                Enable-CsUser -Identity $Using:account.userPrincipalName -RegistrarPool $Using:connectionSettings.RegistrarPool -SipAddress $SipAddress -ErrorAction Stop

                Set-CsUser -Identity $Using:account.userPrincipalName -EnterpriseVoiceEnabled $Using:connectionSettings.EnterpriseVoiceEnabled -ExchangeArchivingPolicy $Using:connectionSettings.ExchangeArchivingPolicy -LineURI $LineUrl -ErrorAction Stop

                Grant-CsDialPlan -Identity $Using:account.userPrincipalName -PolicyName ''  -ErrorAction Stop
                Grant-CsVoicePolicy -Identity $Using:account.userPrincipalName -PolicyName ''  -ErrorAction Stop

                Grant-CsExternalAccessPolicy -Identity $Using:account.userPrincipalName -PolicyName '' -ErrorAction Stop
                Write-Verbose -Verbose "[SkypeForBusiness] Finished enabling user '$($account.userName)'"

                $success = $True
                $auditMessage = "[SkypeForBusiness] Finished enabling user '$($account.userName)'"

            } catch {
                throw $_
            }
        }
        Invoke-SkypeFBCommand -Session $session -Command $command
    } catch [System.Exception] {
        $ex = $_
        $success = $false
        $auditMessage = "[SkypeForBusiness] Could not enable user '$($account.userName)' Error: $ex"
        Write-Error "[SkypeForBusiness] Could not enable user ''$($account.userName)' Error: $ex"
    }
}

#build up result
$result = [PSCustomObject]@{
	Success = $success
	AccountReference = $account_guid
	AuditDetails = $auditMessage
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
