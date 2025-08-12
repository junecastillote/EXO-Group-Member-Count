# Get-EXOGroupMemberCount.ps1

## Overview

`Get-EXOGroupMemberCount.ps1` is a PowerShell script for retrieving member counts and owner details for Exchange Online groups, including:

- Distribution Groups
- Dynamic Distribution Groups
- Microsoft 365 Groups (Unified Groups)

It supports exporting results to CSV or returning them directly in the console.

## Requirements

- PowerShell 7+
- ExchangeOnlineManagement module
- Exchange Online administrative permissions

## Features

- Retrieves group type, primary SMTP address, member count, and Teams-enabled status.
- Optionally resolves owner details.
- Supports pipeline input for multiple groups.
- Can output to both CSV and console.
- Validates Exchange Online connection before processing.

## Usage

```powershell
# Basic usage with console output
./Get-EXOGroupMemberCount.ps1 -Identity "GroupName" -ReturnResult

# Export results to CSV
./Get-EXOGroupMemberCount.ps1 -Identity "GroupName" -OutputCsv "output.csv"

# Pipeline usage with owner resolution
"Group1", "Group2" | ./Get-EXOGroupMemberCount.ps1 -ResolveOwner -OutputCsv "output.csv"
```

## Parameters

- **Identity** *(Mandatory)*: Identity of the group to retrieve information for.
- **ResolveOwner** *(Switch)*: Retrieves the full owner list for the group.
- **OutputCsv** *(String)*: Path to export results to CSV.
- **ReturnResult** *(Switch)*: Returns results to console.

## Script Architecture Overview

```mermaid
flowchart TD
    %% USER INPUT
    subgraph User
        P1[Identity param]
        P2[ResolveOwner switch]
        P3[OutputCsv param]
        P4[ReturnResult switch]
    end

    %% SCRIPT BLOCKS
    subgraph Script
        A[Start Script]
        B[BEGIN block]
        C{EXO Connected?}
        D[Init result list & Defaults]
        E[Prepare output file if needed]
        F[PROCESS block]
        G[Get-Recipient for each Identity]
        H[Identify group type]
        I[Retrieve group details]
        J[Count members]
        K{Resolve owners?}
        L[Get owners info]
        M[Skip owner resolution]
        N[Store in results]
        O[END block]
        P[Export CSV if requested]
        Q[Return results if requested]
        R[Done]
    end

    %% EXCHANGE ONLINE
    subgraph ExchangeOnline
        E1[Get-ConnectionInformation]
        E2[Get-Recipient]
        E3["Get-DistributionGroup / Get-DynamicDistributionGroup / Get-UnifiedGroup"]
        E4["Get-DistributionGroupMember / Get-Recipient (for owners)"]
    end

    %% CONNECTION FLOW
    P1 --> A
    P2 --> A
    P3 --> A
    P4 --> A
    A --> B
    B --> C
    C -->|No| Z[Throw error & Exit]
    C -->|Yes| D
    D --> E
    E --> F

    %% PROCESS FLOW
    F --> G
    G --> H
    H --> I
    I --> J
    J --> K
    K -->|Yes| L
    K -->|No| M
    L --> N
    M --> N
    N --> O
    O --> P
    O --> Q
    P --> R
    Q --> R

    %% DATA FLOW TO/FROM EXO
    B --> E1
    E1 --> B
    G --> E2
    E2 --> G
    I --> E3
    E3 --> I
    J --> E4
    L --> E4
    E4 --> L
```

## Notes

- The script will exit early if not connected to Exchange Online.
- If neither `-OutputCsv` nor `-ReturnResult` is provided, `-ReturnResult` will be enabled by default.
