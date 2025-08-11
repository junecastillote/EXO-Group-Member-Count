[CmdletBinding()]
param (
    [Parameter(Mandatory, ValueFromPipeline )]
    $Identity
)

begin {}
process {

    $teamsEnabled = 'No'
    try {

        $recipientObject = Get-Recipient -Identity $Identity -ErrorAction Stop

        switch ($recipientObject.RecipientTypeDetails) {
            { $_ -in 'MailNonUniversalGroup', 'MailUniversalDistributionGroup', 'MailUniversalSecurityGroup' } {
                $groupObject = Get-DistributionGroup -Identity $recipientObject.Identity
                $groupMemberCount = @(Get-DistributionGroupMember -Identity $recipientObject.Identity -ResultSize Unlimited).Count
                $teamsEnabled = 'N/A'
            }

            { $_ -eq 'DynamicDistributionGroup' } {
                $groupObject = Get-DynamicDistributionGroup -Identity $recipientObject.Identity
                $groupMemberCount = @(Get-Recipient -RecipientPreviewFilter $groupObject.RecipientFilter -ResultSize Unlimited).Count
                $teamsEnabled = 'N/A'
            }

            { $_ -eq 'GroupMailbox' } {
                $groupObject = Get-UnifiedGroup -Identity $recipientObject.Identity
                $groupMemberCount = $groupObject.GroupMemberCount
                if ($groupObject.ResourceProvisioningOptions -contains 'Team') {
                    $teamsEnabled = 'Yes'
                }
            }

            default {
                continue
            }
        }

        $owners = @(
            if ($groupObject.ManagedBy) {
                (
                    $groupObject.ManagedBy | ForEach-Object {
                        $temp = Get-Recipient -Identity $_
                        if ($temp.WindowsLiveId) { $temp.WindowsLiveId }
                        else { $temp.Id }
                    }
                )
            }
        )

        [PSCustomObject]@{
            GroupName    = $groupObject.Name
            GroupEmail   = $groupObject.PrimarySmtpAddress
            GroupType    = $groupObject.RecipientTypeDetails
            TeamsEnabled = $teamsEnabled
            Owners       = $owners -join ", "
            MemberCount  = $groupMemberCount
        }

    }
    catch {
        Write-Error $_.Exception.Message
        continue
    }
}
end {}
