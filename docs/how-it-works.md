# How it works

The script only **reads** data. It changes nothing in Microsoft 365, Power Platform or Azure. Results are written only to the machine that runs it.

## Systems it talks to

| System | What it answers |
|---|---|
| Power Platform licensing service (the data behind PPAC → Licensing → Copilot Studio) | Who used which Copilot Studio agent, which feature, and how many credits |
| Microsoft Graph (Microsoft Entra ID) | User ID → email, display name, department, cost center |
| Azure Cost Management | What was actually billed to the Azure subscription, per billing policy per day |
| Microsoft 365 admin center Copilot Credits report (a CSV you export) | Metered Copilot Chat agents: user × agent × billing policy, as 7- or 30-day totals |

## Step by step

1. **Sign in once.** A device code is shown. Enter it in a browser and sign in as a Power Platform administrator. That one sign-in covers all three online systems, and it is cached so later runs don't ask again.
2. **Read the billing policies.** For each pay-as-you-go policy the script learns:
   - which environments are linked to it
   - which Azure subscription pays for it
   - whether Copilot Studio is switched on for it
3. **Read environment names,** so the report shows names instead of IDs.
4. **Walk through every day in the range:**
   - Ask which users consumed credits that day.
   - For each of those users, ask which agents and features they used, with billed and non-billable credits.
   - Check that each user's agent rows add up to that user's daily total.

   Walking user by user rebuilds the per-agent view that Microsoft closed in September 2026.
5. **Read capacity:** prepaid credits allocated and used, per environment and for the tenant.
6. **Load the Microsoft 365 export,** if you pass one.
7. **Look up users** in Microsoft Graph, 20 per call: email, display name, department, cost center.
8. **Read the actual Azure charges** per billing policy per day, keeping Copilot meters only.
9. **Allocate cost.** Each billing policy's actual charge for the day is split across report rows by billed credits. Microsoft 365 rows use the charge for their 7- or 30-day window instead.
10. **Write the files:** the Excel workbook (formulas, rate as an input cell) and the CSV credits table.

## API calls

| # | System | Call | Purpose |
|---|---|---|---|
| 1 | Microsoft Entra ID | Device-code sign-in (MSAL) | Tokens for the three online systems |
| 2 | Power Platform API | `GET /licensing/billingPolicies` | Policy ↔ environments, Copilot Studio PAYG flag, paying subscription |
| 3 | Power Platform API | `GET /environmentmanagement/environments` | Environment names |
| 4 | Power Platform API | `GET /licensing/entitlements/MCSMessages/users?fromDate=D&toDate=D` | Users with consumption on day D |
| 5 | Power Platform API | `GET /licensing/entitlements/MCSMessages/users/{userId}/resources?fromDate=D&toDate=D` | Agent, feature, billed and non-billable credits for that user and day |
| 6 | Power Platform API | `GET /licensing/environments/{envId}/entitlements`, `GET /licensing/entitlements/MCSMessages` | Prepaid allocation and usage totals |
| 7 | Microsoft Graph | `POST /v1.0/$batch` with `GET /users/{id}?$select=...` | Name, department, cost center |
| 8 | Azure Cost Management | `POST /subscriptions/{id}/providers/Microsoft.CostManagement/query` | Actual daily cost and quantity per billing policy |

- **Power Platform API:** all calls go to `https://api.powerplatform.com` with `api-version=2024-10-01`.
- **Credit names:** `MCSMessages` is the entitlement ID for Copilot Credits. It dates from when credits were called "messages".
- **POST calls are reads:** Graph `$batch` and the Cost Management query use POST, but they don't change anything.

## Terms used in the report

| Term | Meaning |
|---|---|
| Billed Credits | Credits that count against prepaid capacity or pay-as-you-go |
| Non-billable Credits | Zero-rated usage, such as Microsoft 365 Copilot licensed users or developer environments |
| Funding Source | `PAYG`, `Prepaid, then PAYG`, `Prepaid capacity`, `Non-billable`, or `PAYG (Microsoft 365 billing policy)` |
| Estimated Cost | Billed credits × rate (default $0.01 list price) |
| Allocated Azure Cost | The policy's actual Azure charge for the day or window, split by billed credits |

## Data latency

- **Licensing data** is a daily snapshot at 00:00 UTC.
- **Azure charges** can take up to 24 hours to appear.
- **The Microsoft 365 Credits report** is near real time, but only offers 7- or 30-day windows.

## Security notes

- **Cached sign-in:** the cache in `~/.copilot-credits-report/token_cache.bin` holds a refresh token for the account that signed in. Protect it, or delete it after use.
- **Personal data:** reports contain user names and departments. Store and share them according to your organization's privacy policy.
- **Keep reports out of git:** `.gitignore` excludes generated reports and the token cache so they don't get committed by accident.
