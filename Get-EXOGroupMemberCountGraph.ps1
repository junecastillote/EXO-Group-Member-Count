# Get-EXOGroupMemberCountGraph.ps1

#Requires -Modules @{ ModuleName="ExchangeOnlineManagement"; ModuleVersion="3.7.0" }

[CmdletBinding()]
param (
    [Parameter(Mandatory, ValueFromPipeline)]
    $Identity,

    [Parameter()]
    [string]$OutputCsv,

    [Parameter()]
    [switch]
    $Append,

    [Parameter()]
    [switch]$ReturnResult
)

begin {
    $scriptStart = [System.Diagnostics.Stopwatch]::StartNew()

    $results = [System.Collections.Generic.List[object]]::new()

    $counter = 0

    $exoConnected = Get-ConnectionInformation
    if (-not $exoConnected) {
        Write-Error "Exchange Online PowerShell is not connected."
        if ($MyInvocation.InvocationName -like '*.ps1') {
            exit 1
        }
        else {
            throw "Exchange Online PowerShell is not connected."
        }
        break
    }

    $graphConnected = Get-MgContext
    if (-not $graphConnected) {
        Write-Error "Microsoft Graph PowerShell is not connected."
        if ($MyInvocation.InvocationName -like '*.ps1') {
            exit 1
        }
        else {
            throw "Microsoft Graph PowerShell is not connected."
        }
        break
    }

    if (-not $OutputCsv -and -not $ReturnResult) {
        Write-Verbose "Both -OutputCsv and -ReturnResult are not used. Enabling -ReturnResult by default."
        $ReturnResult = $true
    }

    if ($OutputCsv -and (Test-Path $OutputCsv) -and -not $Append) {
        Write-Verbose "The output file [$OutputCsv)] already exists and will be overwritten."
        Remove-Item $OutputCsv -Force -Confirm:$false
    }

    if ($OutputCsv -and (Test-Path $OutputCsv) -and $Append) {
        Write-Verbose "The output file [$OutputCsv)] already exists and new results will be appended."
        # Remove-Item $OutputCsv -Force -Confirm:$false
    }

    $acceptedTypes = @(
        'Deserialized.Microsoft.Exchange.Data.Directory.Management.DynamicDistributionGroup',
        'Deserialized.Microsoft.Exchange.Data.Directory.Management.DistributionGroup',
        'Deserialized.Microsoft.Exchange.Data.Directory.Management.UnifiedGroupBase'
    )
}

process {

    try {
        $counter++

        # Get the object typename
        $objTypeName = $Identity.psobject.typenames[0]

        # Validate by typename
        switch ($objTypeName) {

            { $_ -in $acceptedTypes } {
                $recipientObject = $Identity
            }

            'System.String' {
                $recipientObject = Get-Recipient -Identity $Identity -ErrorAction Stop
            }

            default {
                continue
            }
        }

        $objTypeName = $recipientObject.psobject.TypeNames[0]
        $recipientId = $recipientObject.Identity

        Write-Verbose "Processing [#$($counter)] Type: [$($recipientObject.RecipientTypeDetails)] | Name: [$($recipientObject.DisplayName) / $($recipientObject.PrimarySmtpAddress)]"

        switch ($recipientObject.RecipientTypeDetails) {
            { $_ -in 'MailNonUniversalGroup', 'MailUniversalDistributionGroup', 'MailUniversalSecurityGroup' } {

                if ($objTypeName -notin $acceptedTypes) {
                    $groupObject = Get-DistributionGroup -Identity $recipientId -IncludeManagedByWithDisplayNames
                }

                else {
                    $groupObject = $recipientObject
                }

                if ($groupObject.ManagedBy -and !$groupObject.ManagedByWithDisplayName) {
                    Write-Verbose "  -> Getting owner(s)"
                    $groupObject = Get-DistributionGroup -Identity $recipientId -IncludeManagedByWithDisplayNames
                }

                $groupOwners = $groupObject.ManagedByWithDisplayName
                $groupMemberCount = (Get-MgGroupMemberCount -GroupId $groupObject.ExternalDirectoryObjectId -ConsistencyLevel eventual)
                $teamsEnabled = 'N/A'
            }

            { $_ -eq 'DynamicDistributionGroup' } {
                if ($objTypeName -notin $acceptedTypes) {
                    $groupObject = Get-DynamicDistributionGroup -Identity $recipientId -IncludeManagedByWithDisplayNames
                }
                else { $groupObject = $recipientObject }

                if ($groupObject.ManagedBy -and !$groupObject.ManagedByWithDisplayName) {
                    Write-Verbose "  -> Getting owner(s)"
                    $groupObject = Get-DynamicDistributionGroup -Identity $recipientId -IncludeManagedByWithDisplayNames
                }

                $groupOwners = $groupObject.ManagedByWithDisplayName
                $groupMemberCount = (Get-DynamicDistributionGroupMember -Identity $groupObject -ResultSize Unlimited | Measure-Object).Count
                $teamsEnabled = 'N/A'
            }

            { $_ -eq 'GroupMailbox' } {
                if ($objTypeName -notin $acceptedTypes) {
                    $groupObject = Get-UnifiedGroup -Identity $recipientId
                }
                else { $groupObject = $recipientObject }

                $groupMemberCount = $groupObject.GroupMemberCount
                $groupOwners = (Get-UnifiedGroupLinks -Identity $groupObject -LinkType Owners | ForEach-Object {
                        $(
                            if ($_.PrimarySmtpAddress) {
                                "($($_.PrimarySmtpAddress), $($_.DisplayName))"
                            }
                            else {
                                "($($_.WindowsLiveId), $($_.DisplayName))"
                            }
                        )
                    })
                $teamsEnabled = if ($groupObject.ResourceProvisioningOptions -contains 'Team') { 'Yes' } else { 'No' }
            }
            default {
                continue
            }
        }

        $results.Add([PSCustomObject]@{
                GroupName    = $groupObject.DisplayName
                GroupEmail   = $groupObject.PrimarySmtpAddress
                GroupType    = $groupObject.RecipientTypeDetails
                TeamsEnabled = $teamsEnabled
                Owners       = $groupOwners -join ";"
                MemberCount  = $groupMemberCount
            })
    }
    catch {
        Write-Error $_.Exception.Message
        continue
    }
}

end {

    if ($OutputCsv) {
        $results | Export-Csv -NoTypeInformation -Path $OutputCsv -Append:$Append
        Write-Verbose "Results saved to [$((Resolve-Path $OutputCsv).Path)]"
    }

    if ($ReturnResult) {
        $results
    }

    $scriptStart.Stop()

    Write-Verbose ("Total script time: {0:N2} seconds" -f $scriptStart.Elapsed.TotalSeconds)
}
