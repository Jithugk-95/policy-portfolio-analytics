import re
from datetime import date
from io import BytesIO

import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="Policy Portfolio Analytics", page_icon="📊", layout="wide")
st.title("📊 Policy Portfolio & Follow-up Analytics")
st.caption("V1.5 • Optimized search • Paginated tables • Standard master workbook")

# ---------------- Configuration ----------------
ALIASES = {
    "policy_no": ["policy number", "policy no", "policy_number", "policyno"],
    "doc": ["doc", "date of commencement", "commencement date"],
    "plan_code": ["plan code", "plan", "t", "t.1"],
    "policy_term": ["policy term", "term", "t.2"],
    "ppt": ["ppt", "premium paying term", "premium payment term"],
    "mode": ["mode", "payment mode", "premium mode"],
    "sum_assured": ["sa", "sum assured", "sum_assured"],
    "premium": ["premium", "premium amount", "amount due"],
    "fup": ["fup", "next premium due date", "next due date"],
    "holder_name": ["name of the policy holder", "name of policy holder", "policy holder name", "holder name"],
    "dob": ["date of birth", "dob"],
    "policy_status": ["policy status", "status"],
    "custom_tag": ["custom tag", "tag", "group tag"],
}
LABELS = {
    "policy_no": "Policy Number", "doc": "DOC", "plan_code": "Plan Code",
    "policy_term": "Policy Term", "ppt": "PPT", "mode": "Mode",
    "sum_assured": "Sum Assured", "premium": "Premium", "fup": "FUP",
    "holder_name": "Policyholder Name", "dob": "Date of Birth",
    "policy_status": "Policy Status", "custom_tag": "Custom Tag",
}
DITTO_FIELDS = ["doc", "plan_code", "policy_term", "ppt", "mode", "sum_assured", "premium", "fup"]
STATUS_TERMS = [
    "DEATH CLAIM APPLIED", "DEATH CLAIM", "LAPSED", "SURRENDER",
    "SURRENDERED", "MATURED", "PAID UP", "PAID-UP", "CANCELLED",
    "CANCELED", "CLOSED", "INFORCE", "IN FORCE",
]
DATE_FIELDS = {"doc": "doc", "fup": "fup", "dob": "dob"}

def norm(x):
    return re.sub(r"[^a-z0-9]+", " ", str(x).strip().lower()).strip()

def find_column(columns, field):
    normalized = {c: norm(c) for c in columns}
    for alias in ALIASES[field]:
        a = norm(alias)
        for c, n in normalized.items():
            if n == a:
                return c
    if field == "plan_code":
        return next((c for c in columns if str(c).strip().lower() == "t"), None)
    if field == "policy_term":
        return next((c for c in columns if str(c).strip().lower() in ("t.1", "t_1", "t 1")), None)
    return None

def blank(x):
    return pd.isna(x) or str(x).strip().lower() in ("", "nan", "none", "null", "<na>")

def ditto(x):
    return not pd.isna(x) and str(x).strip() == ",,"

def parse_date(value, kind):
    if blank(value) or ditto(value):
        return pd.NaT
    if isinstance(value, pd.Timestamp):
        return value.normalize()
    raw = str(value).strip()
    if re.fullmatch(r"\d+\.0+", raw):
        raw = raw.split(".", 1)[0]
    if kind == "dob":
        m = re.search(r"(\d{1,2}[-/]\d{1,2}[-/]\d{4})", raw)
        if m:
            d = pd.to_datetime(m.group(1), errors="coerce", dayfirst=True)
            return d.normalize() if not pd.isna(d) else pd.NaT
    digits = re.sub(r"\D", "", raw)
    try:
        if kind == "fup":
            if len(digits) == 5:
                month, year = int(digits[0]), int(digits[1:])
            elif len(digits) == 6:
                month, year = int(digits[:2]), int(digits[2:])
            else:
                month = year = None
            if month and 1 <= month <= 12 and 1900 <= year <= 2200:
                return pd.Timestamp(year=year, month=month, day=1)
        elif kind in ("doc", "dob") and len(digits) == 8:
            return pd.Timestamp(year=int(digits[4:]), month=int(digits[2:4]), day=int(digits[:2]))
    except (ValueError, TypeError, OverflowError):
        return pd.NaT
    d = pd.to_datetime(raw, errors="coerce", dayfirst=True)
    return d.normalize() if not pd.isna(d) else pd.NaT

def numeric(series):
    return pd.to_numeric(series.astype(str).str.replace(r"[₹,\s]", "", regex=True), errors="coerce")

def status_in_row(values):
    text = " ".join(str(v).strip().upper() for v in values if not blank(v))
    for term in STATUS_TERMS:
        if term in text:
            return term.title()
    return None

def distinct_join(values):
    seen = []
    for v in values:
        if blank(v) or ditto(v):
            continue
        if isinstance(v, pd.Timestamp):
            s = v.strftime("%d-%m-%Y")
        else:
            s = str(v).strip()
        if s and s not in seen:
            seen.append(s)
    return ", ".join(seen) if seen else pd.NA

def latest_date(values):
    s = pd.to_datetime(pd.Series(values), errors="coerce").dropna()
    return s.max() if len(s) else pd.NaT

@st.cache_data(show_spinner=False, max_entries=3)
def load_master(file_bytes, sheet_name, mapping_pairs):
    # Cache the expensive Excel parse and normalization. Search keystrokes won't re-read Excel.
    mapping = dict(mapping_pairs)
    raw = pd.read_excel(BytesIO(file_bytes), sheet_name=sheet_name, dtype=object)
    raw.columns = [str(c).strip() for c in raw.columns]
    df = pd.DataFrame(index=raw.index)
    for field in ALIASES:
        col = mapping.get(field)
        df[field] = raw[col] if col and col in raw.columns else pd.NA

    df["excel_row"] = raw.index + 2
    # Apply ditto expansion only to the requested policy columns.
    for field in DITTO_FIELDS:
        s = df[field].mask(df[field].apply(ditto), pd.NA)
        df[field] = s.ffill()

    for field, kind in DATE_FIELDS.items():
        df[field] = df[field].apply(lambda v: parse_date(v, kind))
    for field in ["premium", "sum_assured"]:
        df[field] = numeric(df[field])

    for field in ["policy_no", "holder_name", "policy_status"]:
        df[field] = df[field].apply(lambda v: pd.NA if blank(v) or ditto(v) else str(v).strip())

    # Status-only rows set the current status for subsequent records.
    current_status = None
    retained = []
    statuses = []
    for i in range(len(df)):
        row_status = status_in_row(raw.iloc[i].tolist())
        has_policy = not blank(df.iloc[i]["policy_no"])
        has_name = not blank(df.iloc[i]["holder_name"])
        if row_status and not has_policy and not has_name:
            current_status = row_status
            retained.append(False)
            statuses.append(current_status)
        else:
            if row_status:
                current_status = row_status
            retained.append(True)
            statuses.append(df.iloc[i]["policy_status"] if not blank(df.iloc[i]["policy_status"]) else current_status)
    df["policy_status"] = statuses
    df = df.loc[retained].copy()

    meaningful = df.apply(lambda r: any(not blank(v) for v in r.drop(labels=["excel_row"], errors="ignore")), axis=1)
    df = df.loc[meaningful].copy()
    df["record_type"] = "Policy"
    df.loc[df["policy_no"].isna() & df["holder_name"].notna(), "record_type"] = "Follow-up Required"
    df = df[df["policy_no"].notna() | df["holder_name"].notna()].copy()

    # Build a single row per policy number, and one placeholder per name-only follow-up.
    policies = df[df["policy_no"].notna()].copy()
    followups = df[df["policy_no"].isna() & df["holder_name"].notna()].copy()
    output = []

    def aggregate_group(group, key, record_type):
        row = {key: group.iloc[0][key], "record_type": record_type}
        for field in df.columns:
            if field in (key, "record_type", "excel_row"):
                continue
            if field == "fup":
                row[field] = latest_date(group[field])
            elif field == "premium":
                vals = group[field].dropna()
                row[field] = vals.iloc[-1] if len(vals) else pd.NA
            elif field == "sum_assured":
                vals = group[field].dropna()
                row[field] = vals.iloc[-1] if len(vals) else pd.NA
            else:
                row[field] = distinct_join(group[field])
        row["excel_row"] = distinct_join(group["excel_row"])
        return row

    for policy_no, group in policies.groupby("policy_no", sort=False, dropna=False):
        output.append(aggregate_group(group, "policy_no", "Policy"))
    for holder, group in followups.groupby("holder_name", sort=False, dropna=False):
        output.append(aggregate_group(group, "holder_name", "Follow-up Required"))

    grouped = pd.DataFrame(output)
    if grouped.empty:
        return df, grouped
    grouped["fup"] = pd.to_datetime(grouped["fup"], errors="coerce")
    today = pd.Timestamp(date.today()).replace(day=1)
    next_month = today + pd.DateOffset(months=1)
    grouped["due_status"] = "Date unavailable"
    has_fup = grouped["fup"].notna()
    grouped.loc[has_fup & (grouped["fup"] < today), "due_status"] = "Past due month"
    grouped.loc[has_fup & (grouped["fup"] >= today) & (grouped["fup"] <= next_month), "due_status"] = "Due this/next month"
    grouped.loc[has_fup & (grouped["fup"] > next_month), "due_status"] = "Later"
    # Precompute searchable lowercase strings once.
    grouped["search_key"] = (
        grouped["policy_no"].fillna("").astype(str) + " " +
        grouped["holder_name"].fillna("").astype(str)
    ).str.casefold()
    return df, grouped

def reset_policy_search():
    st.session_state["policy_search_input"] = ""
    st.session_state["policy_search_applied"] = ""

def holder_key(value):
    return "" if blank(value) else str(value).strip().casefold()

def show_page_table(frame, key, page_size_default=25):
    if frame.empty:
        st.info("No records to display.")
        return
    c1, c2 = st.columns([1, 3])
    with c1:
        page_size = st.selectbox("Rows per page", [10, 25, 50, 100], index=1, key=f"{key}_size")
    pages = max(1, (len(frame) + page_size - 1) // page_size)
    page_key = f"{key}_page"
    if st.session_state.get(page_key, 1) > pages:
        st.session_state[page_key] = 1
    with c2:
        page = st.number_input("Page", min_value=1, max_value=pages, value=1, step=1, key=page_key)
    start = (int(page) - 1) * page_size
    end = min(start + page_size, len(frame))
    st.caption(f"Showing {start + 1:,}–{end:,} of {len(frame):,} records • Page {page} of {pages}")
    st.dataframe(frame.iloc[start:end], use_container_width=True, hide_index=True)

# ---------------- User interface ----------------
uploaded = st.sidebar.file_uploader("Upload master Excel workbook", type=["xlsx", "xlsm"])
if not uploaded:
    st.info("Upload a master workbook to begin. For best results, use the standard template included with this application.")
    st.markdown("The app supports a single master worksheet, blank spacer rows, `,,` ditto values in selected policy fields, and name-only follow-up placeholders.")
    st.stop()

file_bytes = uploaded.getvalue()
try:
    xl = pd.ExcelFile(BytesIO(file_bytes))
    sheet_name = st.sidebar.selectbox("Master worksheet", xl.sheet_names)
    header_preview = pd.read_excel(xl, sheet_name=sheet_name, nrows=5, dtype=object)
except Exception as exc:
    st.error(f"Unable to inspect workbook: {exc}")
    st.stop()

st.sidebar.caption("Workbook selected. First-time processing may take a few seconds.")
st.sidebar.subheader("Column mapping")
available = ["(Not present)"] + [str(c) for c in header_preview.columns]
mapping = {}
for field in ALIASES:
    found = find_column(header_preview.columns, field)
    default = found if found in available else "(Not present)"
    selected = st.sidebar.selectbox(LABELS.get(field, field), available, index=available.index(default), key=f"map_{field}")
    mapping[field] = None if selected == "(Not present)" else selected

mapping_pairs = tuple((k, v) for k, v in mapping.items())
with st.spinner("Reading and preparing the workbook… This may take a few seconds on the first upload."):
    try:
        source_rows, grouped = load_master(file_bytes, sheet_name, mapping_pairs)
    except Exception as exc:
        st.error(f"Workbook processing failed: {exc}")
        st.stop()

if grouped.empty:
    st.warning("No policy or follow-up records were detected. Please verify the POLICY NUMBER and policyholder name mappings.")
    st.stop()

# Custom tags are maintained in this local Streamlit session and keyed by policyholder.
# All policies belonging to the same named policyholder inherit the same tag.
if "policyholder_tags" not in st.session_state:
    st.session_state["policyholder_tags"] = {}
tag_map = st.session_state["policyholder_tags"]
if "custom_tag" not in grouped.columns:
    grouped["custom_tag"] = ""
if "custom_tag" not in source_rows.columns:
    source_rows["custom_tag"] = ""
grouped["custom_tag"] = grouped.apply(
    lambda r: tag_map.get(holder_key(r["holder_name"]), "") or ("" if blank(r["custom_tag"]) else str(r["custom_tag"])),
    axis=1,
)
source_rows["custom_tag"] = source_rows.apply(
    lambda r: tag_map.get(holder_key(r["holder_name"]), "") or ("" if blank(r["custom_tag"]) else str(r["custom_tag"])),
    axis=1,
)
policies = grouped[grouped["record_type"].eq("Policy")].copy()
followups = grouped[grouped["record_type"].eq("Follow-up Required")].copy()
premium_total = policies["premium"].sum(min_count=1)

st.caption(f"Ready • {len(source_rows):,} usable source rows • {len(policies):,} unique policy numbers • {len(followups):,} follow-up placeholders")
st.subheader("Business overview")
a, b, c, d = st.columns(4)
a.metric("Total Policy Records", f"{len(source_rows):,}", help="Usable source rows after blank spacer and status-heading rows are handled.")
b.metric("Total Policies (Unique policy numbers)", f"{policies['policy_no'].nunique():,}")
c.metric("Total Premium (Unique policies)", f"₹{premium_total:,.2f}" if pd.notna(premium_total) else "—")
d.metric("Follow-up Required", f"{len(followups):,}")
e, f, g = st.columns(3)
e.metric("Due this/next month", f"{int(grouped['due_status'].eq('Due this/next month').sum()):,}")
f.metric("Past due month", f"{int(grouped['due_status'].eq('Past due month').sum()):,}")
g.metric("Missing FUP", f"{int(grouped['fup'].isna().sum()):,}")

tab1, tab2, tab3, tab4 = st.tabs(["Portfolio analytics", "Follow-up tracker", "Policy search & tags", "Data quality & export"])
with tab1:
    p1, p2 = st.columns(2)
    with p1:
        if not policies.empty:
            chart = policies.groupby("plan_code", dropna=False).size().reset_index(name="Policies")
            st.plotly_chart(px.bar(chart, x="plan_code", y="Policies", title="Unique policies by plan code"), use_container_width=True)
    with p2:
        if not policies.empty:
            chart = policies.groupby("mode", dropna=False).size().reset_index(name="Policies")
            st.plotly_chart(px.pie(chart, names="mode", values="Policies", title="Unique policies by premium mode"), use_container_width=True)
    status_chart = policies.assign(policy_status=policies["policy_status"].fillna("Not specified")).groupby("policy_status").size().reset_index(name="Policies")
    if not status_chart.empty:
        st.plotly_chart(px.bar(status_chart, x="policy_status", y="Policies", title="Policy status distribution"), use_container_width=True)
    st.subheader("Unique policy portfolio")
    portfolio_cols = ["policy_no", "holder_name", "policy_status", "plan_code", "policy_term", "ppt", "mode", "sum_assured", "premium", "fup", "due_status"]
    portfolio = policies[[x for x in portfolio_cols if x in policies.columns]].copy()
    portfolio["fup"] = pd.to_datetime(portfolio["fup"], errors="coerce").dt.strftime("%m/%Y").fillna("")
    show_page_table(portfolio, "portfolio")

with tab2:
    st.subheader("Follow-up required")
    followup_cols = ["holder_name", "policy_status", "excel_row", "doc", "plan_code", "policy_term", "ppt", "mode", "premium", "fup", "dob"]
    followup_view = followups[[x for x in followup_cols if x in followups.columns]].copy()
    if "fup" in followup_view:
        followup_view["fup"] = pd.to_datetime(followup_view["fup"], errors="coerce").dt.strftime("%m/%Y").fillna("")
    show_page_table(followup_view, "followup")

with tab3:
    st.subheader("Policyholder search & group tagging")
    st.caption("Search policyholders and manage family/business tags from this single page.")

    # Search controls
    with st.form("policy_search_form", clear_on_submit=False):
        query_col, submit_col = st.columns([4, 1])
        with query_col:
            query_input = st.text_input(
                "Search by policy number, policyholder name or tag",
                placeholder="Enter part of a name, policy number or tag…",
                key="policy_search_input",
            )
        with submit_col:
            st.write("")
            st.write("")
            submitted = st.form_submit_button("Search", use_container_width=True)
    reset_col, info_col = st.columns([1, 4])
    with reset_col:
        st.button("Reset", on_click=reset_policy_search, use_container_width=True)
    with info_col:
        st.caption("Press Enter or click Search. Reset clears the search and returns all records.")

    if "policy_search_applied" not in st.session_state:
        st.session_state["policy_search_applied"] = ""
    if submitted:
        st.session_state["policy_search_applied"] = query_input.strip()

    applied_query = st.session_state["policy_search_applied"]
    matches = grouped.copy()
    if applied_query:
        needle = applied_query.casefold()
        matches = matches[
            matches["search_key"].str.contains(re.escape(needle), regex=True, na=False)
            | matches["custom_tag"].fillna("").astype(str).str.casefold().str.contains(re.escape(needle), regex=True, na=False)
        ].copy()
        st.caption(f"{len(matches):,} matching grouped records for “{applied_query}”.")
    else:
        st.caption(f"Showing all {len(matches):,} grouped records.")

    search_cols = ["policy_no", "holder_name", "custom_tag", "record_type", "policy_status", "excel_row", "doc",
                   "plan_code", "policy_term", "ppt", "mode", "sum_assured", "premium", "fup", "dob", "due_status"]
    search_view = matches[[x for x in search_cols if x in matches.columns]].copy()
    if "fup" in search_view:
        search_view["fup"] = pd.to_datetime(search_view["fup"], errors="coerce").dt.strftime("%m/%Y").fillna("")
    search_view = search_view.rename(columns={"custom_tag": "Custom Tag"})
    show_page_table(search_view, "search")

    st.divider()
    st.subheader("Assign family / business tags")
    st.caption("Select one or more policyholders and assign a group name. Every policy number belonging to each selected policyholder receives that tag.")
    named_holders = sorted(
        [str(x).strip() for x in grouped["holder_name"].dropna().unique() if str(x).strip()],
        key=str.casefold,
    )
    selected_holders = st.multiselect(
        "Select policyholders",
        named_holders,
        help="A policyholder can own multiple policy numbers; all of their policies inherit the assigned tag.",
        key="tag_selected_holders",
    )
    tag_col, apply_col = st.columns([3, 1])
    with tag_col:
        tag_value = st.text_input("Tag / group name", placeholder="e.g. Sharma family", key="tag_value_input")
    with apply_col:
        st.write("")
        st.write("")
        apply_tag = st.button("Apply tag", use_container_width=True, disabled=not selected_holders or not tag_value.strip())
    if apply_tag:
        for holder in selected_holders:
            tag_map[holder_key(holder)] = tag_value.strip()
        st.session_state["policyholder_tags"] = tag_map
        st.success(f"Applied “{tag_value.strip()}” to {len(selected_holders)} policyholder(s).")
        st.rerun()

    if tag_map:
        if st.button("Clear all custom tags"):
            st.session_state["policyholder_tags"] = {}
            st.rerun()

    tagged = grouped[grouped["custom_tag"].fillna("").astype(str).str.strip().ne("")].copy()
    if not tagged.empty:
        st.subheader("Group summary")
        summary = tagged.groupby("custom_tag", dropna=False).agg(
            Policyholders=("holder_name", "nunique"),
            Policies=("policy_no", "nunique"),
            Premium=("premium", "sum"),
        ).reset_index().rename(columns={"custom_tag": "Tag"})
        show_page_table(summary, "tag_summary", 10)

        st.subheader("View a tagged group")
        tag_filter = st.selectbox("Select tag", ["All tags"] + sorted(tagged["custom_tag"].dropna().unique().tolist()))
        tagged_view = tagged if tag_filter == "All tags" else tagged[tagged["custom_tag"] == tag_filter]
        tag_cols = ["custom_tag", "holder_name", "policy_no", "policy_status", "premium", "fup", "due_status"]
        tagged_view = tagged_view[[x for x in tag_cols if x in tagged_view.columns]].copy()
        if "fup" in tagged_view:
            tagged_view["fup"] = pd.to_datetime(tagged_view["fup"], errors="coerce").dt.strftime("%m/%Y").fillna("")
        tagged_view = tagged_view.rename(columns={"custom_tag": "Custom Tag"})
        show_page_table(tagged_view, "tagged")
    else:
        st.info("No tags assigned yet. Select policyholders above and apply a group name.")

with tab4:
    st.subheader("Data quality")
    quality = pd.DataFrame({
        "Check": ["Source rows", "Unique policy numbers", "Follow-up placeholders", "Missing policy number", "Missing policyholder name", "Missing FUP"],
        "Count": [len(source_rows), policies["policy_no"].nunique(), len(followups),
                  int(source_rows["policy_no"].isna().sum()), int(source_rows["holder_name"].isna().sum()),
                  int(grouped["fup"].isna().sum())],
    })
    show_page_table(quality, "quality", 10)
    st.caption("For reliable searching, use one policy per row, explicit column headers, no merged cells, and no `,,` ditto markers where possible.")
    export_cols = ["policy_no", "holder_name", "custom_tag", "record_type", "policy_status", "excel_row", "doc", "plan_code",
                   "policy_term", "ppt", "mode", "sum_assured", "premium", "fup", "dob", "due_status"]
    export = grouped[[x for x in export_cols if x in grouped.columns]].copy()
    export["fup"] = pd.to_datetime(export["fup"], errors="coerce").dt.strftime("%m/%Y").fillna("")
    export = export.rename(columns={"custom_tag": "Custom Tag"})
    st.download_button("⬇️ Download grouped portfolio (CSV)", export.to_csv(index=False).encode("utf-8-sig"),
                       "policy_portfolio.csv", "text/csv")
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        export.to_excel(writer, index=False, sheet_name="Grouped Portfolio")
        source_rows.to_excel(writer, index=False, sheet_name="Processed Source Rows")
        quality.to_excel(writer, index=False, sheet_name="Data Quality")
    st.download_button("⬇️ Download portfolio report (Excel)", buf.getvalue(),
                       "policy_portfolio_report.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.divider()
st.caption("Local processing • Search index cached during session • No database • No permanent workbook storage")
