# Get-EXOGroupMemberCount.ps1

#Requires -Modules @{ ModuleName="ExchangeOnlineManagement"; ModuleVersion="3.7.0" }


[CmdletBinding()]
param (
    [Parameter(Mandatory, ValueFromPipeline)]
    $Identity,

    [Parameter()]
    [switch]$ResolveOwner,

    [Parameter()]
    [string]$OutputCsv,

    [Parameter()]
    [switch]$ReturnResult
)

begin {
    $scriptStart = [System.Diagnostics.Stopwatch]::StartNew()

    $results = [System.Collections.Generic.List[object]]::new()
    $allOwnerIds = [System.Collections.Generic.HashSet[string]]::new()

    $connected = Get-ConnectionInformation
    if (-not $connected) {
        Write-Error "Exchange Online PowerShell is not connected."
        if ($MyInvocation.InvocationName -like '*.ps1') {
            exit 1
        }
        else {
            throw "Exchange Online PowerShell is not connected."
        }
        break
    }

    if (-not $OutputCsv -and -not $ReturnResult) {
        Write-Verbose "Both -OutputCsv and -ReturnResult are not used. Enabling -ReturnResult by default."
        $ReturnResult = $true
    }

    if ($OutputCsv -and (Test-Path $OutputCsv)) {
        Write-Verbose "The output file [$OutputCsv)] already exists and will be overwritten."
        Remove-Item $OutputCsv -Force -Confirm:$false
    }
}

process {
    try {
        $recipientObject = Get-Recipient -Identity $Identity -ErrorAction Stop
        $recipientId = $recipientObject.Identity
        Write-Verbose "Processing [$($recipientObject.DisplayName)] | [$($recipientObject.RecipientTypeDetails)]"

        switch ($recipientObject.RecipientTypeDetails) {
            { $_ -in 'MailNonUniversalGroup', 'MailUniversalDistributionGroup', 'MailUniversalSecurityGroup' } {
                $groupObject = Get-DistributionGroup -Identity $recipientId
                $groupMemberCount = @(Get-DistributionGroupMember -Identity $recipientId -ResultSize Unlimited).Count
                $teamsEnabled = 'N/A'
            }
            { $_ -eq 'DynamicDistributionGroup' } {
                $groupObject = Get-DynamicDistributionGroup -Identity $recipientId
                $groupMemberCount = @(Get-Recipient -RecipientPreviewFilter $groupObject.RecipientFilter -ResultSize Unlimited).Count
                $teamsEnabled = 'N/A'
            }
            { $_ -eq 'GroupMailbox' } {
                $groupObject = Get-UnifiedGroup -Identity $recipientId
                $groupMemberCount = $groupObject.GroupMemberCount
                $teamsEnabled = if ($groupObject.ResourceProvisioningOptions -contains 'Team') { 'Yes' } else { 'No' }
            }
            default {
                continue
            }
        }

        # Collect owners for later resolution if needed
        if ($ResolveOwner -and $groupObject.ManagedBy) {
            foreach ($ownerId in $groupObject.ManagedBy) {
                $null = $allOwnerIds.Add([string]$ownerId)
            }
        }

        $results.Add([PSCustomObject]@{
                GroupName    = $groupObject.Name
                GroupEmail   = $groupObject.PrimarySmtpAddress
                GroupType    = $groupObject.RecipientTypeDetails
                TeamsEnabled = $teamsEnabled
                Owners       = $groupObject.ManagedBy  # Temp placeholder
                MemberCount  = $groupMemberCount
            })
    }
    catch {
        Write-Error $_.Exception.Message
        continue
    }
}

end {
    $ownerResolutionStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $ownerMap = @{}

    # Resolve owners now in a single loop to avoid duplicates
    if ($ResolveOwner -and $allOwnerIds.Count -gt 0) {
        Write-Verbose "Performing owner name resolution..."
        foreach ($ownerId in $allOwnerIds) {
            try {
                $recipient = Get-Recipient -Identity $ownerId -ErrorAction Stop
                $ownerMap[$ownerId] = if ($recipient.WindowsLiveId) { $recipient.WindowsLiveId } else { $recipient.Id }
            }
            catch {
                $ownerMap[$ownerId] = $ownerId
            }
        }
    }
    $ownerResolutionStopwatch.Stop()

    # Replace placeholder Owners in results
    foreach ($item in $results) {
        if ($ResolveOwner -and $item.Owners) {
            $item.Owners = ($item.Owners | ForEach-Object { $ownerMap[$_] }) -join ', '
        }
        elseif ($item.Owners) {
            $item.Owners = $item.Owners -join ', '
        }
        else {
            $item.Owners = ''
        }
    }

    if ($OutputCsv) {
        $results | Export-Csv -NoTypeInformation -Path $OutputCsv
        Write-Verbose "Results saved to [$((Resolve-Path $OutputCsv).Path)]"
    }

    if ($ReturnResult) {
        $results
    }

    $scriptStart.Stop()

    Write-Verbose ("Owner resolution time: {0:N2} seconds" -f $ownerResolutionStopwatch.Elapsed.TotalSeconds)
    Write-Verbose ("Total script time: {0:N2} seconds" -f $scriptStart.Elapsed.TotalSeconds)
}
