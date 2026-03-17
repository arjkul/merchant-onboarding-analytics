# merchant-onboarding-analytics

> Python automation for extracting and analyzing B2B retailer go-live data from Asana — built to surface onboarding patterns across 200+ wholesale merchants at ShipMonk.

---

## Overview

ShipMonk onboards hundreds of B2B merchants per year. Tracking go-live status, integration health, and account tier across all of them manually was a recurring ops burden. This project automates the extraction of merchant onboarding data from the **B2B Retailer Setup** Asana project and produces a structured Excel report for analysis.

---

## What It Does

1. Connects to the Asana API and paginates through all tasks in the B2B Retailer Setup project
2. Extracts structured fields: merchant name, connection type, account tier, B2B vs dropship, warehouse assignment, first order date, integration status
3. Handles Asana's custom field schema (enum fields vs free-text fields)
4. Outputs a clean Excel workbook for stakeholder distribution

---

## Key Findings (from 200+ merchant extract)

| Dimension | Finding |
|---|---|
| Most common connection type | Manually Created Orders |
| Most common account tier | Platinum |
| Owner assignment | Majority of tasks had no assigned owner |
| Warehouse distribution | Skewed toward Nevada and Pennsylvania sites |
| Integration status | ~30% of merchants had incomplete integration records |

---

## Script

```python
import requests
import json
import pandas as pd
from openpyxl import Workbook

# --- Config ---
ASANA_TOKEN = "your_token_here"
PROJECT_GID = "1209184101505347"  # B2B Retailer Setup project
BASE_URL = "https://app.asana.com/api/1.0"

HEADERS = {
    "Authorization": f"Bearer {ASANA_TOKEN}",
    "Accept": "application/json"
}

CUSTOM_FIELD_MAP = {
    "Connection Type": "connection_type",
    "Account Tier": "account_tier",
    "B2B / Dropship": "b2b_dropship",
    "Warehouse": "warehouse",
    "First Order Date": "first_order_date",
    "Integration Status": "integration_status"
}


def get_tasks_page(project_gid, offset=None):
    """Fetch one page of tasks from the Asana project."""
    params = {
        "project": project_gid,
        "opt_fields": "name,assignee.name,custom_fields,completed,created_at",
        "limit": 100
    }
    if offset:
        params["offset"] = offset

    response = requests.get(f"{BASE_URL}/tasks", headers=HEADERS, params=params)
    response.raise_for_status()
    return response.json()


def extract_custom_fields(task):
    """Extract custom field values from a task, handling enum and text types."""
    fields = {}
    for cf in task.get("custom_fields", []):
        field_name = cf.get("name")
        if field_name not in CUSTOM_FIELD_MAP:
            continue

        col = CUSTOM_FIELD_MAP[field_name]

        # Enum fields
        if cf.get("type") == "enum" and cf.get("enum_value"):
            fields[col] = cf["enum_value"]["name"]
        # Text / date fields
        elif cf.get("text_value"):
            fields[col] = cf["text_value"]
        else:
            fields[col] = None

    return fields


def fetch_all_tasks(project_gid):
    """Paginate through all tasks in the project."""
    all_tasks = []
    offset = None

    while True:
        data = get_tasks_page(project_gid, offset)
        tasks = data.get("data", [])
        all_tasks.extend(tasks)

        # Check for next page
        next_page = data.get("next_page")
        if next_page and next_page.get("offset"):
            offset = next_page["offset"]
        else:
            break

    return all_tasks


def build_dataframe(tasks):
    """Convert task list to a structured DataFrame."""
    rows = []
    for task in tasks:
        row = {
            "merchant_name": task.get("name"),
            "assignee": task.get("assignee", {}).get("name") if task.get("assignee") else None,
            "completed": task.get("completed"),
            "created_at": task.get("created_at")
        }
        row.update(extract_custom_fields(task))
        rows.append(row)

    return pd.DataFrame(rows)


def export_to_excel(df, output_path="merchant_onboarding_report.xlsx"):
    """Export DataFrame to formatted Excel workbook."""
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Merchant Data", index=False)

        # Auto-size columns
        ws = writer.sheets["Merchant Data"]
        for col in ws.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 50)

    print(f"Exported {len(df)} merchants to {output_path}")


if __name__ == "__main__":
    print("Fetching merchant onboarding tasks from Asana...")
    tasks = fetch_all_tasks(PROJECT_GID)
    print(f"Retrieved {len(tasks)} tasks")

    df = build_dataframe(tasks)
    export_to_excel(df)
```

---

## Output Schema

| Column | Source | Notes |
|---|---|---|
| `merchant_name` | Task name | Asana task title |
| `assignee` | Task assignee | Often null — key finding |
| `completed` | Task status | Boolean |
| `created_at` | Task metadata | ISO timestamp |
| `connection_type` | Custom field (enum) | e.g. Manually Created Orders, EDI, API |
| `account_tier` | Custom field (enum) | e.g. Platinum, Gold, Silver |
| `b2b_dropship` | Custom field (enum) | B2B or Dropship classification |
| `warehouse` | Custom field (enum) | Primary warehouse assignment |
| `first_order_date` | Custom field (text) | Go-live date |
| `integration_status` | Custom field (text) | Integration completion status |

---

## Requirements

```
requests
pandas
openpyxl
```

Install with: `pip install requests pandas openpyxl`

---

## Related Repos

- [`warehouse-ops-intelligence`](https://github.com/arjkul/warehouse-ops-intelligence) — What happens after merchants go live
- [`pm-data-stack-templates`](https://github.com/arjkul/pm-data-stack-templates) — Onboarding frameworks and PM templates
