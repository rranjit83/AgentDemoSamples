# Artifacts: PAYG Billing Policy Management for Power Platform

Companion artifacts for the blog post **"Herding Clouds: Taming Pay-As-You-Go Billing Policies in Power Platform at Scale"**.

This folder contains everything you need to follow along end-to-end: a bulk-assignment script, an Azure Automation runbook, a test harness, sample webhook data, and the importable Power Automate solution.

---

## Folder Structure

```
artifacts/
├── scripts/
│   ├── bulk-assign-billing-policy.ps1      # Bulk-link environments to billing policies via CSV
│   ├── UnlinkBillingPolicyRunbook.ps1      # Azure Automation runbook (triggered by budget alert)
│   └── TestRunbook.ps1                     # Manual runbook trigger for end-to-end testing
├── samples/
│   └── Webhooktestdata.json                # Simulated Azure Monitor budget alert payload
└── solution/
    └── BillingPolicyManagement_1_0_0_3.zip # Importable Power Automate solution
```

---

## Prerequisites

| Requirement | Details |
|---|---|
| **Azure CLI** | Installed and authenticated (`az login`) |
| **PowerShell 7+** | Required to run the scripts |
| **Azure Automation Account** | With a System-Assigned Managed Identity |
| **Power Platform role** | Power Platform Admin, Global Admin, or Dynamics 365 Admin |
| **Power Automate environment** | To import the solution into |

> **Note on permissions:** The Automation Account's Managed Identity only needs permission to call the Power Automate HTTP trigger (token audience: `https://service.flow.microsoft.com/`). The Power Platform Admin permissions are held by the connection credentials configured inside the Power Automate solution — those connections must be authenticated by an account with Power Platform Admin rights.

---

## Scripts

### `bulk-assign-billing-policy.ps1`

Bulk-assigns Power Platform environments to billing policies from a CSV file. Runs in six stages: verifies Azure CLI login, validates the CSV, resolves billing policies by name, resolves environment IDs by display name (with tenant-wide pagination for large tenants), links each environment to its policy, and writes results back to the CSV.

**CSV format expected:**

```csv
EnvironmentName,EnvironmentID,BillingPolicyName,Status
Sales-Production,,ProductionBillingPolicy,
Marketing-Sandbox,a1b2c3d4-e5f6-...,DevBillingPolicy,
HR-Production,,ProductionBillingPolicy,
```

- `EnvironmentID` is optional — the script resolves it from `EnvironmentName` if blank.
- `Status` is populated by the script after each run (`Succeeded` or `Failed: <reason>`).
- Only **Production** and **Sandbox** environments are eligible. Developer, Trial, and Default types are skipped with a clear status message.

**Usage:**

```powershell
# Preview what would happen — always run this first
.\bulk-assign-billing-policy.ps1 -InputFile ".\environments.csv" -DryRun

# Execute for real
.\bulk-assign-billing-policy.ps1 -InputFile ".\environments.csv"
```

---

### `UnlinkBillingPolicyRunbook.ps1`

An Azure Automation runbook that acts as the bridge between an Azure Budget alert and the Power Automate unlinking flow. Intended to be hosted in an Azure Automation Account and triggered via webhook from an Azure Action Group.

**What it does:**
1. Parses the incoming Azure Monitor Common Alert Schema webhook payload
2. Extracts the subscription ID and resource group from the `alertId` path
3. Authenticates using the Automation Account's Managed Identity (`Connect-AzAccount -Identity`)
4. Acquires a bearer token for the Power Automate service endpoint
5. POSTs the subscription and resource group context to the Power Automate HTTP trigger flow

**Before deploying, update the hardcoded values:**

| Line | Variable | What to replace with |
|---|---|---|
| 27 | `$Url` | The HTTP trigger URL from your imported Power Automate solution (found in the flow's trigger details) |

```powershell
# Line 27 — replace with your own flow trigger URL
$Url = "https://<your-environment>.environment.api.powerplatform.com/powerautomate/..."
```

---

### `TestRunbook.ps1`

Manually triggers the `UnlinkBillingPolicies` runbook with a local test payload file — so you can validate the entire chain end-to-end without waiting for an actual budget breach.

**Before running, update the hardcoded values at the top of the file:**

| Variable | Description |
|---|---|
| `$SubscriptionId` | Your Azure subscription ID |
| `$AutomationAccountName` | Your Automation Account name |
| `$ResourceGroupName` | Resource group hosting the Automation Account |
| `$RunbookName` | Name of the runbook as deployed in the Automation Account |

**Usage:**

```powershell
.\TestRunbook.ps1
```

This triggers the runbook with the payload from `../samples/Webhooktestdata.json`. The runbook parses it, calls Power Automate, and the flow unlinks all environments from the matching billing policy — full end-to-end, no real spend required.

---

## Samples

### `Webhooktestdata.json`

A realistic Azure Monitor Common Alert Schema payload simulating a budget threshold breach. Used by `TestRunbook.ps1` to trigger the runbook manually.

**Simulated scenario:**
- Budget name: `prodbilling`
- Monthly budget: `$2.00`
- Alert threshold: `$1.60` (80%)
- Simulated spend: `$4.00` (200% — someone's flow has been busy)

The `alertId` field in this payload encodes a real subscription ID and resource group (`Azurevnetforpowerplatform`). The runbook extracts these to identify which billing policy to act on. **Update this file** if your test environment uses a different subscription or resource group.

---

## Solution

### `BillingPolicyManagement_1_0_0_3.zip`

An importable Power Automate solution containing:

| Component | Purpose |
|---|---|
| **Custom connector** (Power Platform Billing Policy API) | Calls `GET /licensing/billingPolicies` to list and look up billing policies by name — not available natively in the Admin V2 connector |
| **HTTP endpoint flow** | Entry point for the runbook; receives subscription ID and resource group, resolves the billing policy name, and delegates to the child flow |
| **UnlinkAllEnvironmentsFromBillingPolicy** (child flow) | Finds the billing policy by name, retrieves all linked environments, and calls `RemoveBillingPolicyEnvironment` for each one — returns a full audit log |

**To deploy:**
1. Import the solution into a Power Platform environment where the connection owner has Power Platform Admin rights
2. Authenticate the two custom connectors during import
3. Copy the HTTP trigger URL from the entry point flow
4. Paste that URL into `UnlinkBillingPolicyRunbook.ps1` at line 27

---

## End-to-End Flow

```
environments.csv
       │
       ▼
bulk-assign-billing-policy.ps1
Environments linked to billing policies ✅
       │
       │   (later, when spend threshold is crossed)
       ▼
Azure Budget Alert → Action Group → Automation Account webhook
       │
       ▼
UnlinkBillingPolicyRunbook.ps1
Parses alert → gets token via Managed Identity → calls Power Automate
       │
       ▼
BillingPolicyManagement solution
Finds policy by name → loops environments → unlinks each one
Environments unlinked ✅  Audit log returned
```

To test the right half of this chain at any time, run `TestRunbook.ps1` with `Webhooktestdata.json` — no real budget breach required.
