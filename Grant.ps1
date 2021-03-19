#####################################################
# HelloID-Conn-Target-SkypeFB-OnPremises-Enable
#
# Version: 1.0.0
#####################################################
#Region functions

#Set from - to range 
$From = 100
$To = 900 
$Exclusions = @(103, 106, 125) #Numbers that should not be given out again

#Policy & Telephonenumber settings
$Resistrarpool = 'pss4bfe01.meierijstad.nl'
$CsDialPlanPolicy = 'Gem_Meierijstad'
$CsVoicePolicyPolicy = 'HoofdNummer_National'
$TelPrefix = '+31413381' 
#
function get-availablenumber {
    [cmdletbinding()]
    param(
          [Parameter(Mandatory)]$From,
          [Parameter(Mandatory)]$To,        
          [Parameter(Mandatory)][String[]]$Exclusions
    )
    
    $UsedNumbers = get-aduser -Properties msRTCSIP-Line -filter {msRTCSIP-Line -like 'tel:*'} | Select-Object @{'Name'='msRTCSIP-Line';'Expression'={"$(-Join $_.'msRTCSIP-Line'[-3..-1])"}}
    $Availablenumbers = @($From..$To) | ForEach-Object { $_.ToString("000") }
    foreach ($Exclusion in $Exclusions)
    {
        $Availablenumbers = $Availablenumbers | Where-Object {$_ -ne [string]$Exclusion}
    }
   
    $CompareTable = compare-object $Availablenumbers $UsedNumbers.'msRTCSIP-Line' | Where-object {$_.SideIndicator -eq "<="} | Select-Object -first 1
    $UseNumber = $CompareTable.InputObject

    write-output $UseNumber
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
    userName = $personObject.Accounts.MicrosoftActiveDirectory.sAMAccountName
    externalId = $account_guid
    samAccountName = $personObject.Accounts.MicrosoftActiveDirectory.sAMAccountName
    userPrincipalName = $personObject.Accounts.MicrosoftActiveDirectory.userPrincipalName
}

if(-Not($dryRun -eq $True)) {
    write-verbose -verbose "[SkypeForBusiness] Starting to enable user '$account.userPrincipalName'"
    try {
        ## Check of PrimaryUserAddress gezet is, voor bestaande gebruikers
        $ADUserExtended = Get-ADUser -Identity $account.samAccountName -Properties msRTCSIP-PrimaryUserAddress, msRTCSIP-Line, userPrincipalName

        ## Check if needed fields contain data -> This might not be  necessary when not using line telephones
        if (($ADUserExtended.'msRTCSIP-PrimaryUserAddress' -like 'sip:*') -and (($ADUserExtended.'msRTCSIP-Line' -like 'tel:*')))
        { 

            $Line = $ADUserExtended.'msRTCSIP-Line'
            $SIPAdress = $ADUserExtended.'msRTCSIP-PrimaryUserAddress'
            $UseNumber = $Line.substring($line.Length - 3)
            write-verbose -verbose "[Skype for Business] User already has active Skype user. Using '$Line' and '$SIPadress'"
        } else 
        {
            ## If not filled, user is not already active, so get useable telephonenumber in range of $From to $To
            $UPN = $ADUserExtended.userPrincipalName
            write-verbose -verbose"[Skype for Business] User $UPN not already activated, getting telephonenumber.."
    
            $useNumber = get-availablenumber -from $From -to $To -exclusions $Exclusions

            if($useNumber -eq $null)
            {
                Write-error "[SkypeForBusiness] Could not find available telephonenumber"
            }
        
            $Line = "tel:$TelPrefix$UseNumber;ext=$UseNumber"
            $SIPAdress = "sip:$UPN"

            write-verbose -verbose"Skype for Business: Activating $UPN using '$Line' and '$SIPadress'"
        } 
  
        #With all info needed, enable user on SFB machine
        $securePassword = $($connectionSettings.Password) | ConvertTo-SecureString -AsPlainText -Force
        [pscredential]$credentials = New-Object System.Management.Automation.PSCredential ($($connectionSettings.UserName), $securePassword)
         
        Invoke-Command -Computer $($connectionSettings.ComputerName) -Credential $credentials -Authentication Credssp -ScriptBlock {
                    write-verbose -verbose "[SkypeForBusiness] Finished setting up credentials"
                    Import-Module 'SkypeForBusiness' -Force
                    write-verbose -verbose "[SkypeForBusiness] Finished importing module"

                    write-verbose -verbose  "[SkypeForBusiness] $Using:SIPAdress"

                    Enable-CsUser -Identity $Using:account.userPrincipalName -RegistrarPool $Using:connectionSettings.Registrarpool -SipAddress $Using:SIPAdress -ErrorAction Stop
                    Set-CsUser -Identity $Using:account.userPrincipalName -EnterpriseVoiceEnabled $Using:connectionSettings.EnterpriseVoiceEnabled -ExchangeArchivingPolicy $Using:connectionSettings.ExchangeArchivingPolicy -LineURI $Using:Line -ErrorAction Stop

                    Grant-CsDialPlan -Identity $Using:account.userPrincipalName -PolicyName $Using:CsDialPlanPolicy  -ErrorAction Stop
                    Grant-CsVoicePolicy -Identity $Using:account.userPrincipalName -PolicyName $Using:CsVoicePolicyPolicy  -ErrorAction Stop

                    Grant-CsExternalAccessPolicy -Identity $Using:account.userPrincipalName -PolicyName "Allow Federation+Public+Outside Access" -ErrorAction Stop
                    Write-Verbose -Verbose "[SkypeForBusiness] Finished enabling user"
                }

        ## Set new business phonenumber as telephonenumber
        Set-ADUser -Identity $account.samAccountName -Replace @{telephonenumber="$TelPrefix$UseNumber"}

        $success = $True

        $auditMessage = "[SkypeForBusiness] Finished enabling user '$UPN'"

            } catch {
                throw $_
                $success = $false
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