# Get-EXOGroupMemberCount.ps1

- [Overview](#overview)
- [Requirements](#requirements)
- [Features](#features)
- [Parameters](#parameters)
- [Output](#output)
- [Usage Examples](#usage-examples)
  - [1. Return results to screen](#1-return-results-to-screen)
  - [2. Export results to CSV](#2-export-results-to-csv)
  - [3. Resolve owner names](#3-resolve-owner-names)
  - [4. Multiple groups via pipeline](#4-multiple-groups-via-pipeline)
- [Script Flow Overview](#script-flow-overview)
- [Script Architecture Overview](#script-architecture-overview)
- [Notes](#notes)

## Overview

`Get-EXOGroupMemberCount.ps1` is a PowerShell script for retrieving member counts and owner details for Exchange Online groups, including:

- Distribution Groups
- Dynamic Distribution Groups
- Microsoft 365 Groups (Unified Groups)

It supports exporting results to CSV or returning them directly in the console.

## Requirements

- **PowerShell**: 7.3 or later (recommended), Windows PowerShell 5.1
- **Modules**: [ExchangeOnlineManagement](https://www.powershellgallery.com/packages/ExchangeOnlineManagement) 3.7.0 or later.
- An active **Exchange Online PowerShell session** (via `Connect-ExchangeOnline`)

## Features

- Retrieves group type, primary SMTP address, member count, and Teams-enabled status.
- Optionally resolves owner details.
- Supports pipeline input for multiple groups.
- Can output to both CSV and console.
- Validates Exchange Online connection before processing.

## Parameters

| Parameter       | Type                    | Mandatory | Description                                                                                           |
| --------------- | ----------------------- | --------- | ----------------------------------------------------------------------------------------------------- |
| `-Identity`     | String / Pipeline input | Yes       | One or more group identities (name, alias, SMTP address, GUID, etc.). Accepts pipeline input.         |
| `-ResolveOwner` | Switch                  | No        | If specified, resolves the **ManagedBy** property to user-friendly identifiers (e.g., WindowsLiveId). |
| `-OutputCsv`    | String                  | No        | Path to a CSV file to export results. **If the file exists, it will be overwritten.**                     |
| `-ReturnResult` | Switch                  | No        | Returns the result objects to the pipeline. Enabled automatically if no output method is specified.   |

---

## Output

The script returns or exports objects with the following properties:

| Property       | Description                                                                     |
| -------------- | ------------------------------------------------------------------------------- |
| `GroupName`    | Display name of the group                                                       |
| `GroupEmail`   | Primary SMTP address                                                            |
| `GroupType`    | Recipient type details (e.g., `GroupMailbox`, `MailUniversalDistributionGroup`) |
| `TeamsEnabled` | `Yes` / `No` for Teams-enabled Microsoft 365 Groups, or `N/A` for other types   |
| `Owners`       | Owner(s) of the group (resolved if `-ResolveOwner` is used)                     |
| `MemberCount`  | Number of members in the group                                                  |

---

## Usage Examples

### 1. Return results to screen

```powershell
.\Get-EXOGroupMemberCount.ps1 -Identity "Marketing Team"
```

### 2. Export results to CSV

```powershell
.\Get-EXOGroupMemberCount.ps1 -Identity "Marketing Team" -OutputCsv "C:\Reports\GroupMembers.csv"
```

### 3. Resolve owner names

```PowerShell
.\Get-EXOGroupMemberCount.ps1 -Identity "Marketing Team" -ResolveOwner
```

### 4. Multiple groups via pipeline

```powershell
"Marketing Team","Sales Team" | .\Get-EXOGroupMemberCount.ps1 -ResolveOwner -OutputCsv ".\Groups.csv"
```

---

## Script Flow Overview

```text
+--------------------------+
|   Start Script           |
+--------------------------+
           |
           v
+--------------------------+
| BEGIN block              |
| - Verify EXO connection  |
| - Initialize result list |
| - Handle defaults        |
| - Prepare output file    |
+--------------------------+
           |
           v
+--------------------------+
| PROCESS block            |
| - For each Identity:     |
|   * Get-Recipient        |
|   * Identify group type  |
|   * Get group details    |
|   * Count members        |
|   * Resolve owners (opt) |
|   * Store in results     |
+--------------------------+
           |
           v
+--------------------------+
| END block                |
| - Export to CSV (if set) |
| - Return results (if set)|
+--------------------------+
           |
           v
+--------------------------+
|         Done             |
+--------------------------+

```

---

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
