# Building one report for all Copilot Credit consumption

Copilot Credits now pay for several workloads, and each workload reports at a different level of detail. This guide shows how to combine them into **one Power BI model**. Each source feeds a common **credits table** at whatever detail it has, and Azure provides the **money**.

## 1. Sources

| Workload | Source | Finest detail | How it lands |
|---|---|---|---|
| Copilot Studio agents | `copilot_credits_report.py` (Power Platform licensing API) | Day × environment × agent × user × feature | Scheduled run → CSV |
| Copilot Studio totals, PAYG vs prepaid split, usage with no user | PPAC → Licensing → Copilot Studio → Download report (environment / agent / user) | Day × environment. Agent and user files are period totals. | Monthly manual export |
| Microsoft 365 Copilot Chat and SharePoint agents (metered) | Microsoft 365 admin center → Reports → Usage → Microsoft Copilot → Credits ("Users and agents") | 7-day window × user × agent × billing policy | Weekly export, loaded with `--m365-credits-csv` |
| Copilot Cowork, Work IQ API (usage-based billing) | Microsoft 365 admin center → Copilot → Cost Management → Consumption, or the Viva Insights consumption dashboard | Week or day × user × service × spending policy | Connector or manual export |
| **Money** for all of the above | Azure Cost Management **scheduled exports** (actual cost, plus amortized cost for pre-purchase plans) | Day × billing policy resource × meter, plus tags | Daily export to a storage account |
| People | Microsoft Graph | Department, cost center, manager, license | Already added by the script; extend as needed |
| Agents | Power Platform inventory, or the Dataverse `bot` table | Owner, environment, created date | Optional scheduled pull |

For Cowork, Work IQ, GitHub Copilot and Azure AI Foundry, Microsoft's open-source [ConsumptionCentral](https://github.com/microsoft/ConsumptionCentral-for-Microsoft-Copilot) Power BI template is a good companion. Its documentation notes that Copilot Studio's **per-user** detail has no API. This tool's CSV fills that gap with dated agent × user rows.

## 2. Data model

**Credits table:** one row per source row. This is the CSV the script writes.

| Column | Notes |
|---|---|
| Date, Period Start, Period End, Grain | `Day` for Copilot Studio; `7 days` or `30 days` for the Microsoft 365 export. Date = Period End. |
| Product, Source, Feature | e.g. `Copilot Studio` / `Power Platform licensing API` / `Generative answer` |
| Agent Name, Agent ID, Environment, Environment ID | Environment is blank for Microsoft 365 agents |
| User, User ID, Department, Cost Center | Department and Cost Center come from Microsoft Graph |
| Billing Policy, Billing Policy ID, Funding Source | Billing policy of the environment (Copilot Studio) or of the user (Microsoft 365) |
| Billed Credits, Non-billable Credits, Credits Consumed | Credits Consumed = billed + non-billable |
| Estimated Cost (USD), Allocated Azure Cost (USD) | Estimated = billed × rate. Allocated = the actual Azure charge split by billed credits. |

**Azure cost table:** date, billing policy resource, meter, quantity, actual cost, amortized cost and tags. This comes from the Cost Management export.

**Lookup tables:** Date, User, Agent, Environment, Billing Policy, Product.

## 3. Rules that keep the numbers honest

- **Azure is the source of truth for money.** The credits table explains it. Allocated cost ties back to the invoice, and anything that can't be matched to a row shows as unmatched cost.
- **Never add overlapping windows.** Export the Microsoft 365 report on the same weekday each week, and don't mix 7-day and 30-day rows.
- **Drill to days only where the source is daily** (Copilot Studio). Views that combine products should use weeks or months.
- **Each product chooses its billing policy differently:**
  - Copilot Studio: by the environment's billing policy
  - Microsoft 365 agents: by the user's billing policy (oldest wins if there are several)
  - Cowork: by spending policy
- **Keep zero-rated usage visible.** Usage by licensed users and in developer environments is value delivered at no extra cost, so don't drop it.

## 4. Report pages

1. **Overview:** credits and cost by product and funding source, month to date against budget, trend and forecast.
2. **Chargeback:** by billing or spending policy and by department; actual against allocated cost.
3. **Agents:** top agents, which features drive cost, users per agent, billed against zero-rated.
4. **Users:** top users, licensed against unlicensed, outliers (for example more than 2,000 credits in 30 days).
5. **Detail:** Date | Agent | User | Billing Policy | Credits | Cost, with Product and Grain.
6. **Reconciliation and freshness:** Azure against the report per policy, unmatched cost, last load per source.
7. **Governance:** agents or policies near their limits, and setup problems, for example a billing policy without Copilot Studio enabled.

## 5. Power BI starter

**Load every credits CSV from a folder.** If a day was exported twice, only the newest file's rows are kept. For SharePoint, replace `Folder.Files` with `SharePoint.Files("https://<tenant>.sharepoint.com/sites/<site>", [ApiVersion = 15])` and filter on `[Folder Path]`.

```m
let
    FolderPath = "C:\CopilotCredits",
    Files = Folder.Files(FolderPath),
    CreditFiles = Table.SelectRows(Files, each Text.StartsWith([Name], "copilot_credits_") and [Extension] = ".csv"),
    Newest = Table.Buffer(Table.Sort(CreditFiles, {{"Date modified", Order.Descending}})),
    Parsed = Table.AddColumn(Newest, "Data", each Table.PromoteHeaders(
        Csv.Document([Content], [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
        [PromoteAllScalars = true])),
    Combined = Table.Buffer(Table.Combine(Parsed[Data])),
    Deduped = Table.Distinct(Combined, {"Source", "Grain", "Period Start", "Period End", "Product", "Agent ID",
        "User", "Feature", "Billing Policy"}),
    Typed = Table.TransformColumnTypes(Deduped, {
        {"Date", type date}, {"Period Start", type date}, {"Period End", type date},
        {"Credits Consumed", type number}, {"Estimated Cost (USD)", type number},
        {"Allocated Azure Cost (USD)", type number}, {"Billed Credits", type number},
        {"Non-billable Credits", type number}}, "en-US")
in
    Typed
```

**Measures.** Name the query `Credits`. Allocation is already calculated in the CSV, so the measures stay simple:

```dax
Billed Credits = SUM ( Credits[Billed Credits] )
Non-billable Credits = SUM ( Credits[Non-billable Credits] )
Credits Consumed = [Billed Credits] + [Non-billable Credits]
Estimated Cost (USD) = SUM ( Credits[Estimated Cost (USD)] )
Allocated Azure Cost (USD) = SUM ( Credits[Allocated Azure Cost (USD)] )
Zero-rated Share = DIVIDE ( [Non-billable Credits], [Credits Consumed] )
```

Add a Date table related to `Credits[Date]`. Use week or month on visuals that mix products.

## 6. Fallback when the API isn't available

If Microsoft restricts the user-level endpoints, load the PPAC **per-user** download (`EntitlementConsumptionTenantPerUserDetailsReport_MCSMessages*.csv`) into the same credits table, mapped as follows:

| Credits table column | PPAC per-user column |
|---|---|
| Grain | `Period` (the file covers month to date) |
| User, User ID | `User Email`, `User Id` |
| Agent Name, Agent ID | `Agent Name`, `Agent Id` |
| Billed Credits | `Billable credit used` |
| Non-billable Credits | `Credits used` − `Billable credit used` |
| Product, Source | `Copilot Studio`, `PPAC per-user download` |

These column names come from observed exports, as documented by ConsumptionCentral; check them against your own file. You lose the daily detail. Export monthly on the same day so the periods don't overlap.

## 7. Constraints to plan for

- **Identity:** the licensing endpoints worked only with a delegated, signed-in admin in testing. Plan for a dedicated admin account with a cached sign-in, or an admin who runs the report weekly.
- **Manual exports:** the Microsoft 365 Credits report and the PPAC downloads have no APIs, so they stay manual. A Power Automate flow can move the files into the landing folder.
- **Cost separation:** Azure shows Copilot Studio, Cowork and Work IQ under one service. Use separate subscriptions or resource groups for each billing method, or Azure tags, to separate them on the invoice.
