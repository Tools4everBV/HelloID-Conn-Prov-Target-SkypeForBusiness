#####################################################
# HelloID-Conn-Target-SkypeFB-OnPremises-Disable
#
# Version: 1.0.0
#####################################################
#Region functions

#Initialize default properties
$connectionSettings = ConvertFrom-Json $configuration
$personObject = ConvertFrom-Json $person

$auditMessage = "for person " + $personObject.DisplayName
$success = $False
$account_guid = New-Guid

#Change mapping here
$account = [PSCustomObject]@{
    displayName = $personObject.DisplayName
    userName = $personObject.Accounts.MicrosoftActiveDirectory.sAMAccountName
    externalId = $account_guid
    samAccountName = $personObject.Accounts.MicrosoftActiveDirectory.sAMAccountName
    userPrincipalName = $personObject.Accounts.MicrosoftActiveDirectory.userPrincipalName
}

if(-Not($dryRun -eq $True)) { 
    write-verbose -verbose "[SkypeForBusiness] Starting to disable user '$account.userPrincipalName'"
    try {
       #Disable user on SFB machine
       $securePassword = $($connectionSettings.Password) | ConvertTo-SecureString -AsPlainText -Force
       [pscredential]$credentials = New-Object System.Management.Automation.PSCredential ($($connectionSettings.UserName), $securePassword)
        
       Invoke-Command -Computer $($connectionSettings.ComputerName) -Credential $credentials -Authentication Credssp -ScriptBlock {
                write-verbose -verbose "[SkypeForBusiness] Finished setting up connection"
                Import-Module 'SkypeForBusiness' -Force
                write-verbose -verbose "[SkypeForBusiness] Finished importing module"

                Import-module "D:\Scripts\Remove-RgsMember.ps1" -Force

                $upn = $Using:account.userPrincipalName
                Remove-RgsMember -User sip:$upn -AllGroups -Verbose
                Disable-CsUser -Identity $upn      
        }

        write-verbose -verbose "[SkypeForBusiness] Finished disabling user '$account.userPrincipalName'"
        $auditMessage = "[SkypeForBusiness] Finished disabling user '$account.userPrincipalName'"
        set-aduser -identity $account.samAccountName -clear @("telephonenumber")
        $success = $True
        
    } catch {
        $ex = $_
        $success = $false
        $auditMessage = "[SkypeForBusiness] Could not disable user '$account.userPrincipalName' Error: $ex"
        Write-Error "[SkypeForBusiness] Could not disable user '$account.userPrincipalName' Error: $ex"
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