# Policy Portfolio & Follow-up Analytics
## Application Behaviour Runbook

**Application:** Policy Portfolio & Follow-up Analytics  
**Current version:** V1.7  
**Application type:** Local, single-user Streamlit web application  
**Primary input:** Excel master workbook (`.xlsx` or `.xlsm`)  
**Database:** None; workbook data is processed in memory  
**Default local URL:** http://localhost:8501

---

# 1. Purpose and scope

This application helps a single user analyse a policyholder master workbook and maintain a searchable policy portfolio.

The application is intended to support:

- Policy-level portfolio reporting.
- Policyholder search when one person owns multiple policies.
- Identification of follow-up placeholders where a name is available but a policy number is missing.
- Premium, plan, payment-mode and policy-status analysis.
- FUP (next premium due month) tracking.
- Family or business grouping using custom tags.
- Paginated tables and downloadable CSV/Excel reports.

The application is a reporting and analysis tool. It is not a policy administration system, payment collection system, insurer system of record or workflow notification engine.

## 1.1 Current limitations

- Single-user, local usage only.
- No database or permanent application-side storage.
- No authentication, role-based access control or audit history.
- No automatic WhatsApp, SMS, email or reminder notifications.
- No confirmed payment/collection tracking unless a suitable payment field and explicit business rule are added in a future version.
- Custom tags assigned through the UI are session-based. Export the processed workbook to retain them, or maintain the `CUSTOM TAG` column in the source workbook.
- The current tag model assigns one tag per policyholder. Multiple tags per policyholder are not supported in this version.

---

# 2. Technology stack

| Component | Technology | Purpose |
|---|---|---|
| User interface | Streamlit | Local web interface |
| Language | Python | Application logic |
| Excel reading/writing | pandas, openpyxl | Workbook processing and exports |
| Charts | Plotly | Portfolio visualisations |
| Data storage | In-memory pandas DataFrames | Temporary processing; no database |

A database installation is not required.

---

# 3. Project structure

Expected extracted folder:

```text
policy_portfolio_analytics_v1_7/
├── app.py
├── requirements.txt
├── README.md
└── run_windows.bat
```

| File | Description |
|---|---|
| `app.py` | Main Streamlit application and business logic |
| `requirements.txt` | Python package dependencies |
| `README.md` | This application runbook |
| `run_windows.bat` | Optional Windows launcher; requires `.venv` to exist |

Do not delete `app.py` or `requirements.txt`. Keep this README with the project so that future maintainers can understand the expected behaviour.

---

# 4. System prerequisites

## 4.1 Supported setup

- Windows 10 or Windows 11.
- Python 3.10 or newer.
- Internet access during the initial package installation.
- A modern browser such as Chrome or Edge.
- An Excel workbook in `.xlsx` or `.xlsm` format.

The application is designed for a small master workbook (for example, a few thousand policy rows). Performance depends on the machine, workbook size, Excel formatting and the number of columns.

## 4.2 Verify Python

Open Command Prompt and run:

```cmd
py --version
python --version
```

If neither command works, install Python from:

https://www.python.org/downloads/

During installation, enable **Add Python to PATH**.

---

# 5. Installation and first-time setup

1. Download and extract the application ZIP.
2. Open Command Prompt.
3. Change directory to the extracted project folder. Example:

```cmd
cd /d D:\my_project\policy_portfolio_analytics_v1_7
```

4. Create a virtual environment:

```cmd
py -m venv .venv
```

5. Activate it:

```cmd
.venv\Scripts\activate
```

The prompt should now start with `(.venv)`.

6. Upgrade pip and install dependencies:

```cmd
python -m pip install --upgrade pip
pip install -r requirements.txt
```

7. Start the application:

```cmd
streamlit run app.py
```

8. Open the local URL shown in the terminal, normally:

http://localhost:8501

## 5.1 Subsequent launches

The virtual environment and dependencies need to be created/installed only once per project folder.

For later sessions:

```cmd
cd /d D:\my_project\policy_portfolio_analytics_v1_7
.venv\Scripts\activate
streamlit run app.py
```

Alternatively, double-click `run_windows.bat` after completing the first-time setup.

## 5.2 Stop the application

Return to the Command Prompt window running Streamlit and press:

```text
Ctrl + C
```

Close the browser tab after the server has stopped.

---

# 6. Input workbook specification

## 6.1 Workbook selection

- Upload one `.xlsx` or `.xlsm` master workbook.
- Select the worksheet containing the master policy data.
- The application displays column-mapping controls in the sidebar.
- Verify the mappings before relying on the results.
- The workbook is processed in memory; the application does not intentionally save the uploaded source workbook to disk.

## 6.2 Main reporting fields

The core fields are:

| Logical field | Common source header | Purpose |
|---|---|---|
| Policy number | `POLICY NUMBER`, `POLICY NO` | Identifies an individual policy |
| DOC | `DOC` | Date of commencement |
| Plan code | First `T` or `PLAN CODE` | Policy plan code |
| Policy term | Second `T` or `POLICY TERM` | Selected policy term |
| PPT | `PPT` | Premium paying term |
| Mode | `MODE` | Premium payment frequency |
| Sum assured | `SA` | Sum assured |
| Premium | `PREMIUM` | Premium value |
| FUP | `FUP` | Next premium due month |
| Policyholder name | `NAME OF THE POLICY HOLDER` | Customer/policyholder name |
| Date of birth | `DATE OF BIRTH`, `DOB` | Date of birth |
| Policy status | `POLICY STATUS`, `STATUS` | Status classification, where supplied |
| Custom tag | `CUSTOM TAG`, `TAG` | Optional family/business grouping label |

The application also supports legacy headers such as duplicate `T` columns. Confirm that the first T maps to Plan Code and the second T maps to Policy Term.

## 6.3 Recommended workbook format

For reliable processing, use the separate `Policy_Master_Template.xlsx` template when available.

Recommended rules:

1. Keep one policy per row.
2. Use a unique, explicit header for each column.
3. Do not merge cells.
4. Avoid blank spacer rows.
5. Repeat values instead of using `,,` ditto markers.
6. Put status in a dedicated `POLICY STATUS` column for each applicable policy.
7. Put a follow-up name in `NAME OF THE POLICY HOLDER` and leave policy number blank when the policy number is not yet known.
8. Keep date and numeric values consistent.
9. Keep the optional `CUSTOM TAG` value consistent for all policies of the same policyholder.
10. Do not add unrelated title rows above the header row.

## 6.4 Sensitive extra columns

The original workbook may contain mobile numbers, Aadhaar, PAN, bank information, email, address, nominee and other personal information.

Only map and process fields required for the current reporting use case. Do not include sensitive columns in reports unless there is a clear business requirement and appropriate authorization.

The current application focuses on the reporting fields described above; it is not designed as a repository for all personal information in the original workbook.

---

# 7. Data preparation and transformation rules

This section documents how source rows are interpreted.

## 7.1 Blank rows

Completely empty rows and blank spacer rows are excluded from the usable source records.

A row containing a policyholder name or policy number is not treated as a blank row.

## 7.2 Ditto marker (` , , `)

Legacy workbooks may use `,,` to mean “copy the value from the preceding row”.

The application forward-fills `,,` only for these fields:

- DOC
- First T / Plan Code
- Second T / Policy Term
- PPT
- MODE
- SA
- PREMIUM
- FUP

The application does not intentionally forward-fill policy number, policyholder name, DOB, nominee, mobile, address, PAN, Aadhaar or bank information.

**Important:** Forward-fill assumes the ditto marker refers to the immediately preceding applicable row. If a workbook uses `,,` across section boundaries, verify the result carefully. The recommended template avoids ditto markers entirely.

## 7.3 Date interpretation

| Field | Example source value | Intended interpretation |
|---|---|---|
| DOC | `12092013` | 12 September 2013 |
| FUP | `122026` | December 2026 |
| FUP | `42026` | April 2026 |
| DOB | `19-05-1995 / SC` | 19 May 1995 |

FUP is displayed as `MM/YYYY`, without a day.

DOC and DOB are displayed as day-month-year where successfully parsed.

Use actual Excel date cells or consistent text formats. Review dates that appear blank because an input value could not be parsed.

## 7.4 Policy rows

A row with a policy number is treated as a policy record, provided the row is retained as meaningful data.

## 7.5 Follow-up placeholders

A row with a policyholder name but no policy number is classified as:

`Follow-up Required`

This is a placeholder for a policyholder whose policy details need investigation or follow-up. It is not counted as a numbered policy.

Name-only placeholder rows are grouped by policyholder name in the grouped results. If two different people have exactly the same name, the current name-based grouping can combine them. Use an additional approved unique customer identifier in a future enhancement if this distinction is required.

## 7.6 Policyholder with multiple policies

A policyholder may own multiple policy numbers.

The application keeps each distinct policy number as a separate policy record. It does not collapse all policies belonging to the same person into one policy row.

The same policyholder can therefore appear in multiple rows in the policy search and grouped portfolio.

## 7.7 Repeated policy numbers

Repeated source rows with the same policy number are grouped into one policy-level record.

Current aggregation behaviour:

| Field | Grouping behaviour |
|---|---|
| Policy number | One grouped record per policy number |
| Policyholder name | Distinct values combined |
| FUP | Latest valid date retained |
| Premium | Last non-empty value in source worksheet order |
| Sum assured | Last non-empty value in source worksheet order |
| Other mapped text/date fields | Distinct values combined where present |
| Source row | Source Excel row numbers combined |

If a policy's premium or sum assured changes over time, the current implementation uses the last non-empty value in source worksheet order. It does not determine whether that change is a correction, renewal, endorsement or a historical snapshot.

## 7.8 Policy status headings

The legacy workbook may contain status-only heading rows, for example:

- DEATH CLAIM APPLIED
- LAPSED
- SURRENDER
- SURRENDERED
- MATURED
- PAID UP
- CANCELLED
- CLOSED
- INFORCE / IN FORCE

The current legacy handling attempts to detect recognized status text in rows without a policy number or policyholder name and carry that status forward to subsequent records.

This is a best-effort compatibility rule. It assumes the status heading applies to subsequent records until another status heading appears.

**Preferred format:** Use a dedicated `POLICY STATUS` column on every policy row. Do not rely on merged cells or status-only headings for future workbooks.

---

# 8. Application pages and behaviour

## 8.1 Business overview

The overview displays high-level portfolio indicators such as:

- Total Policy Records: usable source rows retained by the application.
- Total Policies (Unique policy numbers): distinct policy numbers after grouping.
- Total Premium (Unique policies): sum of the grouped premium values, with each policy number counted once.
- Follow-up Required: grouped name-only placeholder records.
- Due this/next month: grouped records whose FUP falls in the current or next month.
- Past due month: grouped records whose FUP is earlier than the current month.
- Missing FUP: grouped records without a valid FUP.

These are reporting calculations, not confirmed payment collection figures.

## 8.2 Portfolio analytics

Charts provide a view of:

- Unique policies by plan code.
- Unique policies by premium mode.
- Policy status distribution.

Charts depend on the availability and correctness of mapped source fields.

## 8.3 Follow-up tracker

Shows name-only records classified as Follow-up Required.

Use this view to identify policyholders for whom policy details are incomplete.

## 8.4 Policy search & tags

This page combines search and custom grouping.

### Search

- Search by partial policyholder name.
- Search by partial policy number.
- Search by custom tag.
- Click Search or press Enter to apply the search.
- Click Reset to clear the applied search.
- Results are paginated.

Search is performed against the grouped policy records. A policyholder with multiple policy numbers may return multiple rows.

### Tags

- Select one or more policyholders.
- Enter a tag such as `Sharma family` or `Aggarwal unit`.
- Apply the tag.
- All policy records for the selected policyholder name inherit that tag.
- Review the group summary and filter the tagged records.

### Tag limitations

- Current UI tag assignment is session-based.
- Restarting the application or starting a new session may lose tags assigned only in the UI.
- Download the processed Excel report after applying tags to retain a copy.
- The `CUSTOM TAG` source column can be used to carry tags into a later upload.
- The current model supports one tag per policyholder. Applying a new tag replaces the previous session tag for that policyholder.
- The association is based on policyholder name, so identical names can be ambiguous.

## 8.5 Data quality & export

Provides basic quality counts and downloadable outputs.

Available outputs:

- Grouped portfolio CSV.
- Grouped portfolio Excel report.
- Processed source rows worksheet.
- Data quality worksheet.

Exports may contain sensitive policyholder information. Store and share them only through approved business channels.

---

# 9. Pagination and search performance

## 9.1 Pagination

Tables provide page-size choices:

- 10 rows
- 25 rows
- 50 rows
- 100 rows

The page indicator shows the displayed row range and total record count.

Pagination limits the number of rows rendered in the browser at a time. It does not reduce the amount of source data processed during initial workbook loading.

## 9.2 Caching

The workbook loading and normalization function uses Streamlit's data cache.

The intention is to avoid re-reading and re-normalizing the Excel file for each search interaction when the uploaded file, worksheet and column mappings are unchanged.

A first upload or changed workbook/mapping requires processing again.

## 9.3 Expected loading behaviour

- During initial processing, a spinner should display.
- After processing, the application shows usable source rows, unique policy numbers and follow-up placeholders.
- Search should operate on the already prepared in-memory data.

If the application appears stuck, check the Command Prompt window for errors and note whether CPU usage or the terminal output changes.

---

# 10. Business calculation definitions

| Metric | Definition |
|---|---|
| Total Policy Records | Number of retained source rows after blank/status-heading handling |
| Total Policies | Number of distinct policy numbers after grouping |
| Total Premium | Sum of one grouped premium value per unique policy number |
| Follow-up Required | Number of grouped name-only placeholders |
| Due this/next month | FUP in current month or next month |
| Past due month | FUP earlier than current month |
| Missing FUP | Grouped records without a parsed FUP |

### Premium caveat

The current premium aggregation uses the last non-empty premium value in worksheet order for a repeated policy number. The dashboard does not sum every repeated source row for that policy.

Confirm this is the intended business definition before using the premium KPI for financial reconciliation.

### Due-date caveat

FUP is a month/year field. Due buckets are based on the first day of the FUP month, not on an exact payment due day.

No payment status is inferred from FUP or last transaction date.

---

# 11. Troubleshooting

## 11.1 `The system cannot find the path specified`

Likely cause: `.venv` does not exist in the current folder or the command is being run from a different directory.

Run:

```cmd
dir
py -m venv .venv
.venv\Scripts\activate
```

## 11.2 `py is not recognized` or `python is not recognized`

Python may not be installed or may not be added to PATH. Install Python and reopen Command Prompt.

## 11.3 `No module named streamlit`

Activate the virtual environment and run:

```cmd
pip install -r requirements.txt
```

## 11.4 Search returns no records

- Confirm the correct worksheet is selected.
- Verify the Policy Number and Policyholder Name mappings.
- Check spelling and spaces in the source name.
- Try a shorter partial name or policy number.
- Click Search or press Enter.
- Click Reset to clear an earlier search.

## 11.5 Duplicate T columns mapped incorrectly

Excel readers often rename duplicate headers to `T` and `T.1`. Confirm:

- First T = Plan Code.
- Second T = Policy Term.

If the workbook has renamed headers, use the sidebar mapping controls.

## 11.6 Dates appear blank or incorrect

Check that:

- DOC is a valid date or `ddMMyyyy`.
- FUP is a valid month/year or supported compact `MMyyyy` value.
- DOB uses a consistent date format.
- Ditto values are used only where the previous row contains the intended value.

## 11.7 Application is slow

- Close other large applications and retry.
- Use the standard template with only required reporting columns.
- Avoid unnecessary formatting, formulas and merged cells in the source workbook.
- Check the terminal for Python errors.
- Confirm the workbook has not changed during upload.
- If performance remains poor, capture the terminal error and approximate workbook row/column count for investigation.

## 11.8 Port 8501 is already in use

Start Streamlit on another local port:

```cmd
streamlit run app.py --server.port 8502
```

Then open http://localhost:8502.

## 11.9 Application stops

Run the application again from the project directory:

```cmd
.venv\Scripts\activate
streamlit run app.py
```

---

# 12. Privacy and security runbook

The workbook can contain highly sensitive personal and financial information.

- Use only a trusted computer and approved local account.
- Do not upload real customer data to public websites or unapproved cloud services.
- Do not expose Streamlit to the public internet.
- Do not use `--server.address 0.0.0.0` unless a properly approved and secured network deployment is designed.
- Restrict access to the project folder and exported reports.
- Avoid including Aadhaar, PAN, bank, mobile, email and address fields unless required.
- Use anonymized records when asking for technical support.
- Follow the organization's data retention and privacy policies.

Local hosting does not automatically mean the data is encrypted. The application does not currently provide encryption at rest, user authentication or audit logging.

---

# 13. Backup and version management

Recommended practice:

1. Keep each application release in a separate versioned folder.
2. Do not overwrite a working version until the new version is tested.
3. Keep a protected backup of the original master workbook.
4. Export the grouped portfolio after making tag changes.
5. Keep the README updated whenever business logic or user-visible behaviour changes.
6. Use Git for source-code version history if the project is maintained in a repository.
7. Never commit real policyholder workbooks, exports, passwords or secrets to a public Git repository.

Suggested folders:

```text
D:\PolicyAnalytics\
├── releases\
│   ├── policy_portfolio_analytics_v1_6\
│   └── policy_portfolio_analytics_v1_7\
├── input\
│   └── Policy_Master_Template.xlsx
├── exports\
└── documentation\
    └── README.md
```

Do not put real customer data in a source-code repository.

---

# 14. Current feature inventory

| Feature | Current behaviour |
|---|---|
| Single master worksheet upload | Supported |
| Excel `.xlsx` / `.xlsm` | Supported |
| Column mapping | Supported |
| Blank spacer row handling | Supported |
| Legacy `,,` ditto expansion | Supported for configured columns |
| Compact DOC/FUP date parsing | Supported |
| Name-only follow-up placeholders | Supported |
| Multiple policies per policyholder | Supported |
| Repeated policy number grouping | Supported |
| Latest FUP selection | Supported |
| Search by name/policy number/tag | Supported |
| Search and Reset buttons | Supported |
| Enter-to-search | Supported |
| Pagination | Supported |
| Family/business tags | Supported; one tag per holder |
| Group summary | Supported |
| CSV and Excel exports | Supported |
| Database persistence | Not available |
| Persistent tag database | Not available |
| Authentication / user roles | Not available |
| Audit trail | Not available |
| Automated reminders | Not available |
| Multi-user access | Not available |

---

# 15. Change log

## V1.7
- Combined Policy Search and Tags & Family Groups into one page.
- Retained search, reset, Enter-to-search and pagination.
- Added tag assignment, group summary and tagged-record filtering on the combined page.

## V1.6
- Added policyholder-level custom tags.
- Added family/business group summary.
- Added tag search and tag export.
- Added Search and Reset controls with Enter submission.

## V1.5
- Added cached workbook preparation and search index.
- Added pagination to tables.
- Added standardized master workbook template.

## V1.4
- Introduced single-sheet master workbook processing.
- Added follow-up placeholder classification.
- Added legacy ditto handling and policy status heading recognition.
- Added portfolio and follow-up analytics.

## V1.3
- Changed FUP display to `MM/YYYY`.

## V1.2
- Updated unique-policy premium calculation and overview metrics.

## V1.1
- Added unique policy grouping and date parsing improvements.

---

# 16. Future enhancement backlog

Potential improvements, subject to business confirmation:

1. Policyholder 360° view showing all policies for a selected person.
2. Persistent tag storage using a local database or a dedicated tag master workbook.
3. Multiple tags per policyholder.
4. Tag editing, reassignment and removal.
5. Customer-level unique identifier to distinguish people with identical names.
6. Follow-up notes, owner, status and next follow-up date.
7. Policy status dashboard with validated status rules.
8. Advanced filters and saved searches.
9. Premium trend and portfolio segmentation.
10. Automated backup, audit logging and authentication if multi-user access is introduced.

---

# 17. Operational checklist

## Start of use
- [ ] Open the correct version folder.
- [ ] Activate `.venv`.
- [ ] Start Streamlit.
- [ ] Upload the approved master workbook.
- [ ] Select the correct worksheet.
- [ ] Verify column mappings.
- [ ] Check total rows, unique policies and follow-up counts.

## Before exporting
- [ ] Review search results.
- [ ] Review follow-up placeholders.
- [ ] Review missing FUP values.
- [ ] Review policy status values.
- [ ] Verify tags and group summaries.
- [ ] Confirm premium calculation assumptions.
- [ ] Save exports in an approved location.

## End of use
- [ ] Download any required reports.
- [ ] Protect or close exported files.
- [ ] Stop Streamlit with Ctrl+C.
- [ ] Close the browser tab.

---

**Document ownership:** Application project documentation  
**Update this runbook whenever application behaviour, calculations, data rules, installation steps or security assumptions change.**
