# Get-EXOGroupMemberCount.ps1

- [Overview](#overview)
- [Requirements](#requirements)
- [Features](#features)
- [Parameters](#parameters)
- [Output](#output)
- [Usage Examples](#usage-examples)
  - [1. Return results to screen](#1-return-results-to-screen)
  - [2. Export results to CSV](#2-export-results-to-csv)
  - [3. Append results to a new or existing CSV](#3-append-results-to-a-new-or-existing-csv)
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
- **Modules**:
  - [ExchangeOnlineManagement](https://www.powershellgallery.com/packages/ExchangeOnlineManagement) 3.7.0 or later.
  - [Microsoft Graph PowerShell SDK](https://www.powershellgallery.com/packages/Microsoft.Graph)
- An active **Exchange Online PowerShell session** (via `Connect-ExchangeOnline`)
- An active **Microsoft Graph PowerShell session** (via `Connect-MgGraph`)

## Features

- Retrieves group type, primary SMTP address, member count, and Teams-enabled status.
- Supports pipeline input for multiple groups.
- Can output to both CSV and console.
- Validates Exchange Online and Microsoft Graph connections before processing.

## Parameters

| Parameter       | Type                    | Mandatory | Description                                                                                         |
| --------------- | ----------------------- | --------- | --------------------------------------------------------------------------------------------------- |
| `-Identity`     | String / Pipeline input | Yes       | One or more group identities (name, alias, SMTP address, GUID, etc.). Accepts pipeline input.       |
| `-OutputCsv`    | String                  | No        | Path to a CSV file to export results. **If the file exists, it will be overwritten.**               |
| `-Append`       | Switch                  | No        | Appends the output to the output CSV file instead of overwriting it.                                |
| `-ReturnResult` | Switch                  | No        | Returns the result objects to the pipeline. Enabled automatically if no output method is specified. |

---

## Output

The script returns or exports objects with the following properties:

| Property       | Description                                                                     |
| -------------- | ------------------------------------------------------------------------------- |
| `GroupName`    | Display name of the group                                                       |
| `GroupEmail`   | Primary SMTP address                                                            |
| `GroupType`    | Recipient type details (e.g., `GroupMailbox`, `MailUniversalDistributionGroup`) |
| `TeamsEnabled` | `Yes` / `No` for Teams-enabled Microsoft 365 Groups, or `N/A` for other types   |
| `Owners`       | Owner(s) of the group.                                                          |
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

### 3. Append results to a new or existing CSV

```PowerShell
.\Get-EXOGroupMemberCount.ps1 -Identity "Marketing Team" -OutputCsv "C:\Reports\GroupMembers.csv" -Append
```

### 4. Multiple groups via pipeline

```powershell
"Marketing Team","Sales Team" | .\Get-EXOGroupMemberCount.ps1 -OutputCsv ".\Groups.csv"
```

---

## Script Flow Overview

```text
+---------------------------------------------------------------+
|                  Get-EXOGroupMemberCountGraph.ps1             |
+---------------------------------------------------------------+
| Parameters: Identity (req), OutputCsv, Append, ReturnResult   |
+---------------------------------------------------------------+
        |
        v
+---------------------------+
| BEGIN block               |
+---------------------------+
| Start stopwatch           |
| Init results list, counter|
|                           |
| Check EXO connection      |
|   ├─ No → Error + exit    |
|   └─ Yes → Continue       |
| Check Graph connection    |
|   ├─ No → Error + exit    |
|   └─ Yes → Continue       |
|                           |
| If no OutputCsv & no      |
| ReturnResult → set        |
| ReturnResult=$true        |
|                           |
| If OutputCsv exists:      |
|   ├─ Append=$false →      |
|       Overwrite file      |
|   └─ Append=$true →       |
|       Append to file      |
|                           |
| Define acceptedTypes[]    |
+---------------------------+
        |
        v
+---------------------------+
| PROCESS block (per item)  |
+---------------------------+
| counter++                 |
| Determine $objTypeName    |
| switch($objTypeName):     |
|   ├─ In acceptedTypes →   |
|         $recipientObject=Identity
|   ├─ "System.String" →    |
|         Get-Recipient     |
|   └─ Else → continue      |
|                           |
| Get $recipientId, log info|
|                           |
| switch($RecipientTypeDetails):
|   ├─ Mail*Group ----------+
|   |   Ensure full groupObj|
|   |   Get owners          |
|   |   Get member count    |
|   |   TeamsEnabled = N/A  |
|   +-----------------------+
|   ├─ DynamicDistribution -+
|   |   Ensure full groupObj|
|   |   Get owners          |
|   |   Get member count    |
|   |   TeamsEnabled = N/A  |
|   +-----------------------+
|   ├─ GroupMailbox --------+
|   |   Ensure full groupObj|
|   |   Get owners          |
|   |   Get member count    |
|   |   TeamsEnabled = Yes/No
|   +-----------------------+
|   └─ Else → continue      |
|                           |
| Add result object to list |
| On error → Write-Error &  |
| continue                  |
+---------------------------+
        |
        v
+---------------------------+
| END block                 |
+---------------------------+
| If OutputCsv → Export-CSV |
| If ReturnResult → output  |
| Stop stopwatch            |
| Verbose: total time       |
+---------------------------+

```

---

## Script Architecture Overview

```mermaid
flowchart TD
    A[Start script] --> B[BEGIN block]
    B --> B1[Start stopwatch]
    B1 --> B2[Init results list and counter]
    B2 --> C{EXO connected}
    C -- No --> C1[Write-Error and exit or throw]
    C -- Yes --> D{Graph connected}
    D -- No --> D1[Write-Error and exit or throw]
    D -- Yes --> E[Set ReturnResult true if no OutputCsv and no ReturnResult]
    E --> F{OutputCsv exists}
    F -- No --> J[Define acceptedTypes list]
    F -- Yes --> F1{Append switch}
    F1 -- False --> H[Overwrite file]
    F1 -- True --> I[Append to file]
    H --> J
    I --> J

    J --> K{Process each identity}
    K --> K1[Increment counter]
    K1 --> K2{Type in acceptedTypes}
    K2 -- Yes --> K3[Set recipientObject to identity]
    K2 -- String --> K4[Get-Recipient]
    K2 -- Other --> K5[Skip item]

    K3 --> L[Get recipientId and log]
    K4 --> L
    L --> M{RecipientTypeDetails}

    M -- Mail group --> M1[Ensure full group object]
    M1 --> M2[Get owners]
    M2 --> M3[Get member count using Graph]
    M3 --> M4[Teams enabled NA]
    M4 --> N[Add result object]

    M -- Dynamic distribution group --> Dg1[Ensure full group object]
    Dg1 --> Dg2[Get owners]
    Dg2 --> Dg3[Get member count using Get DynamicDistributionGroupMember]
    Dg3 --> Dg4[Teams enabled NA]
    Dg4 --> N

    M -- Group mailbox --> Gm1[Ensure full group object]
    Gm1 --> Gm2[Get owners using Get UnifiedGroupLinks]
    Gm2 --> Gm3[Get member count from property]
    Gm3 --> Gm4[Teams enabled Yes or No]
    Gm4 --> N

    M -- Other --> K5

    N --> O[Errors are logged and processing continues]
    O --> P[After last item]

    P --> Q[END block]
    Q --> Q1{OutputCsv specified}
    Q1 -- Yes --> Q2[Export CSV]
    Q1 -- No --> Q3[Skip CSV export]
    Q2 --> Q4{ReturnResult specified}
    Q3 --> Q4
    Q4 -- Yes --> Q5[Output results]
    Q4 -- No --> Q6[Skip result output]
    Q5 --> Q7[Stop stopwatch and write total time]
    Q6 --> Q7
    Q7 --> R[End script]
```

## Notes

- The script will exit early if not connected to Exchange Online.
- If neither `-OutputCsv` nor `-ReturnResult` is provided, `-ReturnResult` will be enabled by default.
