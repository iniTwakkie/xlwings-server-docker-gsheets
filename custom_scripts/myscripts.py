from xlwings.server import script
import requests
import datetime as dt
import xlwings as xw
import re
from urllib.parse import urlencode
import uuid

SETTINGS_SHEET_NAME = "Settings"
TB_SHEET_NAME = "TB"
COA_SHEET_NAME = "COA"
XERO_REDIRECT_URI = "https://local-xlwings.danienell.com/xeroconnect"
XERO_SCOPES = " ".join([
    "offline_access",
    "accounting.reports.read",
    "accounting.settings.read",
])

@script
def hello_world(book: xw.Book):
    sheet = book.sheets.active
    cell = sheet["H1"]
    if cell.value == "Hello xlwings!":
        cell.value = "Bye xlwings!"
    else:
        cell.value = "Hello xlwings!"

@script
def yellow(book: xw.Book):
    """
    Highlights the currently selected cells in Excel in yellow.
    """
    try:
        # Get the active sheet
        sheet = book.sheets.active

        # Get the currently selected range
        selected_range = book.app.selection

        # Set the background color to yellow (RGB: 255, 255, 0)
        selected_range.color = "#FFFF00"

        print("Selected cells highlighted in yellow!")
        
    except Exception as e:
        print(f"Error: {e}")

    # Return the following response
    return book.json()

def _read_settings(sheet):
    # Read key/value pairs from Settings: A=key, B=value
    rng = sheet.range("A1").expand()
    values = rng.value
    settings = {}

    if values is None:
        return settings

    # Normalize to a list of rows
    if not isinstance(values[0], (list, tuple)):
        values = [values]

    for row in values:
        if not row:
            continue

        raw_key = row[0]
        if raw_key is None or raw_key == "":
            # skip completely empty key cells
            continue

        key = str(raw_key).strip()

        # Skip the header row "Key"
        if key.lower() == "key":
            continue

        value = row[1] if len(row) > 1 else None
        settings[key] = value

    return settings



def _write_setting(sheet, key, value):
    """Write a key/value pair into Settings.
    If key exists in column A, update its B cell; otherwise append a new row.
    """
    key_clean = str(key).strip()

    rng = sheet.range("A1").expand()
    values = rng.value

    # If sheet is empty, just write to A1/B1
    if values is None:
        sheet["A1"].value = key_clean
        sheet["B1"].value = value
        return

    num_rows = rng.rows.count  # e.g. 33 rows → indices 0..32 on rng[...]

    for i in range(num_rows):
        # rng[row_index, col_index] is 0-based relative to A1
        raw_key = rng[i, 0].value  # col 0 = column A
        if raw_key is None or raw_key == "":
            continue

        try:
            cell_key = str(raw_key).strip()
        except Exception as e:
            print(f"_write_setting: skipping row {i+1} key={raw_key!r} (error {e})")
            continue

        if cell_key == key_clean:
            # col 1 = column B
            rng[i, 1].value = value
            return

    # Append new row at the bottom (1-based Excel row numbering here)
    next_row = num_rows + 1
    sheet[f"A{next_row}"].value = key_clean
    sheet[f"B{next_row}"].value = value


def _get_access_token_from_refresh(settings, settings_sheet):
    client_id = settings["xero_client_id"]
    client_secret = settings["xero_client_secret"]
    refresh_token = settings["xero_refresh_token"]

    resp = requests.post(
        "https://identity.xero.com/connect/token",
        data={
            "grant_type": "refresh_token",
            "refresh_token": refresh_token,
        },
        auth=(client_id, client_secret),
        headers={"Accept": "application/json"},
        timeout=30,
    )
    resp.raise_for_status()
    token_data = resp.json()

    access_token = token_data["access_token"]
    new_refresh = token_data.get("refresh_token")

    # Xero rotates refresh tokens, so store the new one if present
    if new_refresh and new_refresh != refresh_token:
        _write_setting(settings_sheet, "xero_refresh_token", new_refresh)

    return access_token


def _fetch_trial_balance(access_token, tenant_id, as_of_date=None):
    url = "https://api.xero.com/api.xro/2.0/Reports/TrialBalance"
    params = {}
    if as_of_date:
        params["date"] = as_of_date  # YYYY-MM-DD

    resp = requests.get(
        url,
        headers={
            "Authorization": f"Bearer {access_token}",
            "xero-tenant-id": tenant_id,
            "Accept": "application/json",
        },
        params=params,
        timeout=30,
    )
    resp.raise_for_status()
    return resp.json()


def _flatten_trial_balance(tb_json, code_to_guid):
    """
    Flatten Xero TrialBalance JSON into rows:

    Columns:
      account_id   (GUID from Accounts API)
      account_code (numeric code)
      section
      account
      YTD Debit
      YTD Credit
    """
    reports = tb_json.get("Reports", [])
    if not reports:
        return []

    report = reports[0]
    rows_out = []

    for row in report.get("Rows", []):
        if row.get("RowType") != "Section":
            continue

        section_name = row.get("Title", "") or ""

        for inner in row.get("Rows", []):
            if inner.get("RowType") != "Row":
                continue

            cells = inner.get("Cells", []) or []
            if not cells:
                continue

            # -----------------------
            # Account name
            # -----------------------
            account_name = cells[0].get("Value", "") or ""

            # -----------------------
            # Account code from Attributes or name
            # -----------------------
            account_code = ""
            attrs = cells[0].get("Attributes", []) or []
            for attr in attrs:
                name = attr.get("Name")
                val = attr.get("Value", "")
                if name == "AccountCode":
                    account_code = val

            # Fallback: parse code from "Name (1234)"
            if not account_code and isinstance(account_name, str):
                m = re.search(r"\(([^()]+)\)\s*$", account_name)
                if m:
                    account_code = m.group(1).strip()

            # -----------------------
            # GUID via mapping from Accounts API
            # -----------------------
            account_guid = ""
            if account_code:
                account_guid = code_to_guid.get(str(account_code).strip(), "")

            # -----------------------
            # YTD Debit / Credit = last two cell values
            # -----------------------
            ytd_debit = ""
            ytd_credit = ""
            if len(cells) >= 2:
                ytd_credit = cells[-1].get("Value", "")
                ytd_debit = cells[-2].get("Value", "")

            rows_out.append([
                account_guid,
                account_code,
                section_name,
                account_name,
                ytd_debit,
                ytd_credit,
            ])

    return rows_out


def _get_tb_meta(tb_json, fallback_date):
    """
    Extract organisation name and report date from the Xero TrialBalance JSON.
    Falls back to the provided fallback_date if needed.
    """
    org_name = ""
    report_date = fallback_date

    reports = tb_json.get("Reports", [])
    if not reports:
        return org_name, report_date

    report = reports[0]

    # 1) Use explicit ReportDate if present
    rd = report.get("ReportDate")
    if rd:
        report_date = rd

    # 2) Look into ReportTitles for org name and "As at ..." date
    titles = report.get("ReportTitles", []) or []
    if len(titles) >= 2:
        org_name = titles[1] or org_name

    for t in titles:
        if isinstance(t, str) and "As at" in t:
            # e.g. "As at 31 March 2010" -> "31 March 2010"
            parts = t.split("As at", 1)
            if len(parts) == 2:
                report_date = parts[1].strip() or report_date
            break

    return org_name, report_date


def _fetch_accounts(access_token, tenant_id):
    resp = requests.get(
        "https://api.xero.com/api.xro/2.0/Accounts",
        headers={
            "Authorization": f"Bearer {access_token}",
            "xero-tenant-id": tenant_id,
            "Accept": "application/json",
        },
        timeout=30,
    )
    resp.raise_for_status()
    return resp.json()


def _build_code_to_guid_map(accounts_json):
    mapping = {}
    for acc in accounts_json.get("Accounts", []):
        code = str(acc.get("Code", "")).strip()
        acc_id = acc.get("AccountID", "")
        if code and acc_id:
            mapping[code] = acc_id
    return mapping

@script
def xero_trial_balance(book: xw.Book):
    """
    xlwings Server entrypoint.

    Called via:
      runPython(getServerUrl() + '/xero_trial_balance')
    from the workbook that holds Settings and TB sheet.
    """
    settings_sheet = book.sheets[SETTINGS_SHEET_NAME]
    settings = _read_settings(settings_sheet)

    tenant_id = settings["xero_tenant_id"]

    # As-of date we REQUEST from Xero (from Settings)
    tb_date_raw = settings.get("xero_tb_date")
    if isinstance(tb_date_raw, dt.date):
        as_of_date = tb_date_raw.strftime("%Y-%m-%d")
    elif isinstance(tb_date_raw, str) and tb_date_raw:
        as_of_date = tb_date_raw
    else:
        as_of_date = dt.date.today().strftime("%Y-%m-%d")

    # 1) get fresh access token using client id + secret + refresh token
    access_token = _get_access_token_from_refresh(settings, settings_sheet)

    # 2) get Accounts to build code → GUID map
    accounts_json = _fetch_accounts(access_token, tenant_id)
    code_to_guid = _build_code_to_guid_map(accounts_json)

    # 3) call Xero TB for THIS workbook's tenant
    tb_json = _fetch_trial_balance(access_token, tenant_id, as_of_date)

    # 4) derive org name + report date from the TB JSON (for cross-checking)
    org_name, report_date = _get_tb_meta(tb_json, as_of_date)

    # 5) flatten into desired columns using the mapping
    data_rows = _flatten_trial_balance(tb_json, code_to_guid)

    # 6) write to TB sheet in THIS workbook
    tb_sheet = book.sheets[TB_SHEET_NAME]
    tb_sheet.range("A1:F2000").clear_contents()

    # Row 1: org name + report date FROM XERO TB RESPONSE
    title_bits = []
    if org_name:
        title_bits.append(org_name)
    title_bits.append(f"Trial Balance as of {report_date}")
    title_text = " - ".join(title_bits)
    tb_sheet["A1"].value = title_text

    # Row 2: fixed headings
    tb_sheet["A2"].value = [
        "account_id",    # GUID
        "account_code",  # numeric code
        "section",
        "account",
        "YTD Debit",
        "YTD Credit"
    ]

    # Data rows start at row 3
    if data_rows:
        tb_sheet["A3"].value = data_rows

    return book.json()


def _fetch_organisation_name(access_token, tenant_id):
    """Return the Organisation Name for the given tenant."""
    resp = requests.get(
        "https://api.xero.com/api.xro/2.0/Organisation",
        headers={
            "Authorization": f"Bearer {access_token}",
            "xero-tenant-id": tenant_id,
            "Accept": "application/json",
        },
        timeout=30,
    )
    resp.raise_for_status()
    data = resp.json()
    orgs = data.get("Organisations", []) or data.get("organisations", [])
    if not orgs:
        return ""
    return orgs[0].get("Name", "")

@script
def xero_coa(book: xw.Book):
    """
    Downloads the Xero Chart of Accounts for the current tenant
    and writes it into the 'COA' sheet:

      Row 1: Organisation name
      Row 2: AccountID | Code | Name | Status | Class
      Row 3+: Data
    """
    settings_sheet = book.sheets[SETTINGS_SHEET_NAME]
    settings = _read_settings(settings_sheet)

    tenant_id = settings["xero_tenant_id"]

    # 1) get fresh access token
    access_token = _get_access_token_from_refresh(settings, settings_sheet)

    # 2) fetch organisation name
    org_name = _fetch_organisation_name(access_token, tenant_id)

    # 3) fetch accounts
    accounts_json = _fetch_accounts(access_token, tenant_id)
    accounts = accounts_json.get("Accounts", []) or []

    # 4) get or create COA sheet
    try:
        coa_sheet = book.sheets[COA_SHEET_NAME]
    except Exception:
        coa_sheet = book.sheets.add(COA_SHEET_NAME)

    # Clear only a reasonable area (avoid 10M-cell issue in Google Sheets)
    coa_sheet.range("A1:E2000").clear_contents()

    # Row 1: org name
    coa_sheet["A1"].value = f"{org_name} - Chart Of Accounts"

    # Row 2: headings
    coa_sheet["A2"].value = ["AccountID", "Code", "Name", "Status", "Class"]

    # Row 3+: data
    rows = []
    for acc in accounts:
        rows.append([
            acc.get("AccountID", ""),
            acc.get("Code", ""),
            acc.get("Name", ""),
            acc.get("Status", ""),
            acc.get("Class", ""),
        ])

    if rows:
        coa_sheet["A3"].value = rows

    return book.json()

@script
def fivetran_start_sync(book: xw.Book):
    """
    Trigger a manual Fivetran sync and write a simple status to Main_Summary!G4
    """
    settings_sheet = book.sheets[SETTINGS_SHEET_NAME]
    settings = _read_settings(settings_sheet)

    connection_id = settings.get("FivetranConnectionID")
    api_key = settings.get("FivetranAPIKey")
    api_secret = settings.get("FivetranAPISecret")

    sheet = book.sheets["Main_Summary"]

    if not connection_id or not api_key or not api_secret:
        sheet["G4"].value = "Missing Fivetran settings"
        return book.json()

    url = f"https://api.fivetran.com/v1/connections/{connection_id}/sync"
    payload = {"force": True}  # matches the tutorial (you can set False if you prefer)

    headers = {
        "Accept": "application/json;version=2",
        "Content-Type": "application/json",
    }

    try:
        # Let requests handle Basic Auth: no manual base64
        resp = requests.post(url, json=payload, headers=headers,
                             auth=(api_key, api_secret), timeout=30)
        resp.raise_for_status()
        data = resp.json()
        print("Fivetran sync response:", data)

        code = (data.get("code") or "").lower()
        if code == "success":
            sync_state = "Syncing"
        else:
            sync_state = f"Not Syncing ({code or 'unknown'})"

    except requests.HTTPError as exc:
        print("Error triggering Fivetran sync:", exc, getattr(exc.response, "text", ""))
        sync_state = f"Error: {exc}"
    except Exception as e:
        print("Unexpected error in fivetran_start_sync:", e)
        sync_state = f"Error: {e}"

    sheet["G4"].value = sync_state
    return book.json()

@script
def fivetran_status(book: xw.Book):
    """
    Fetches the Fivetran connection details and writes the sync status
    into Main_Summary!G4.
    """
    settings_sheet = book.sheets[SETTINGS_SHEET_NAME]
    settings = _read_settings(settings_sheet)

    connector_id = settings.get("FivetranConnectorID")
    base64_key = settings.get("FivetranBase64APIkey")

    sheet = book.sheets["Main_Summary"]

    if not connector_id or not base64_key:
        sheet["G4"].value = "Missing Fivetran settings"
        return book.json()

    # ✅ Again, /connections not /connectors
    url = f"https://api.fivetran.com/v1/connections/{connector_id}"

    headers = {
        "Accept": "application/json;version=2",
        "Authorization": f"Basic {base64_key}",
    }

    try:
        resp = requests.get(url, headers=headers, timeout=30)
        resp.raise_for_status()
        data = resp.json()
        # print("Fivetran status response:", data)

        status = data.get("data", {}).get("status", {})
        sync_state_raw = status.get("sync_state", "unknown")

        # Map into something friendly for the sheet
        if sync_state_raw == "scheduled":
            sync_state = "Not Syncing"
        elif sync_state_raw == "syncing":
            sync_state = "Syncing"
        else:
            sync_state = sync_state_raw or "unknown"

    except requests.HTTPError as exc:
        body = getattr(exc.response, "text", "")
        print(f"Error fetching Fivetran status: {exc} | body={body}")
        sync_state = f"Error: {exc}"

    except Exception as e:
        print(f"Unexpected error in fivetran_status: {e}")
        sync_state = f"Error: {e}"

    sheet["G4"].value = sync_state
    return book.json()

def _make_xero_auth_url(client_id: str, state: str) -> str:
    params = {
        "response_type": "code",
        "client_id": client_id,
        "redirect_uri": XERO_REDIRECT_URI,
        "scope": XERO_SCOPES,
        "state": state,
    }
    return "https://login.xero.com/identity/connect/authorize?" + urlencode(params)


@script
def xero_connect_xw(book: xw.Book):
    """
    Start Xero OAuth flow:
      - reads xero_client_id / xero_client_secret from Settings
      - generates & stores xero_state
      - builds the Xero auth URL
      - stores xero_auth_url in Settings
      - returns the workbook JSON (no extra payload)
    """
    settings_sheet = book.sheets[SETTINGS_SHEET_NAME]
    settings = _read_settings(settings_sheet)

    client_id = settings.get("xero_client_id")
    client_secret = settings.get("xero_client_secret")

    # No sheet writes, no custom JSON: just log and bail out
    if not client_id or not client_secret:
        print("xero_connect_xw: Missing xero_client_id or xero_client_secret")
        return book.json()

    # Generate new state and save it
    state = uuid.uuid4().hex
    _write_setting(settings_sheet, "xero_state", state)

    # Build the OAuth URL
    auth_url = _make_xero_auth_url(client_id, state)

    # Save the auth_url into Settings
    _write_setting(settings_sheet, "xero_auth_url", auth_url)

    # Just return the standard workbook JSON (no arguments allowed here)
    return book.json()

@script
def xero_choose_tenant(book: xw.Book, tenant_name: str):
    """
    Select a Xero tenant for this workbook:
      - reads xero_connections_json from Settings
      - finds the matching tenantName
      - writes xero_tenant_id and xero_tenant_name into Settings
    """
    settings_sheet = book.sheets[SETTINGS_SHEET_NAME]
    settings = _read_settings(settings_sheet)

    connections_raw = settings.get("xero_connections_json")
    if not connections_raw:
        print("xero_choose_tenant: No xero_connections_json found in Settings.")
        return book.json()

    try:
        connections = json.loads(connections_raw)
    except Exception as e:
        print(f"xero_choose_tenant: Failed to parse xero_connections_json: {e}")
        return book.json()

    match = None
    for c in connections:
        if c.get("tenantName") == tenant_name:
            match = c
            break

    if not match:
        print(f"xero_choose_tenant: tenantName {tenant_name!r} not found.")
        return book.json()

    _write_setting(settings_sheet, "xero_tenant_id", match.get("tenantId"))
    _write_setting(settings_sheet, "xero_tenant_name", match.get("tenantName"))

    print(f"xero_choose_tenant: Selected tenant {match.get('tenantName')} ({match.get('tenantId')})")
    return book.json()

