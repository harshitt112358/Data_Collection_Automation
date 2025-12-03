# app.py
# Streamlit app: Multi-Function Data Collection OFT Generator
# - Landing page: pick a function (ER&D, Supply Chain, Procurement, Manufacturing)
# - ER&D uses approved wording (exact body text preserved)
# - Non-ER&D functions use identical UI/logic with placeholder "lorem ipsum" bodies
# - Includes hardening: Jinja auto-escape, Excel dtype=str, single Outlook COM session reuse,
#   recipient normalization & dedup, validation, per-row status.

from __future__ import annotations
import io
import os
import re
import sys
import zipfile
import tempfile
from datetime import datetime
from typing import List, Dict

import pandas as pd
import streamlit as st
from jinja2 import Environment, select_autoescape

# -------------------------------------
# Environment: require Windows + Outlook (pywin32)
# -------------------------------------
WINDOWS = sys.platform.startswith("win")
try:
    import win32com.client  # type: ignore
    import pythoncom  # type: ignore
    HAS_WIN32 = True
except Exception:
    HAS_WIN32 = False

# -------------------------------------
# Jinja (safe HTML auto-escape)
# -------------------------------------
_JINJA_ENV = Environment(
    autoescape=select_autoescape(enabled_extensions=(), default=True)
)

def render_text(template_str: str, context: dict) -> str:
    tmpl = _JINJA_ENV.from_string(template_str)
    return tmpl.render(**context)

# -------------------------------------
# Helpers
# -------------------------------------
def sanitize_filename(name: str) -> str:
    bad = '<>:"/\\|?*\n\r\t'
    for ch in bad:
        name = name.replace(ch, "-")
    return " ".join(name.split()).strip()

def _split_recipients(s: str) -> List[str]:
    s = str(s or "").replace(",", ";")  # allow comma-separated inputs from Excel
    return [p.strip() for p in s.split(";") if p.strip()]

def build_cc(*chunks: str) -> str:
    """Normalize + merge semicolon lists, de-dup (case-insensitive), keep order."""
    parts = []
    for s in chunks:
        parts.extend(_split_recipients(s))
    seen = set()
    out = []
    for p in parts:
        key = p.lower()
        if key in seen:
            continue
        seen.add(key)
        out.append(p)
    return "; ".join(out)

def dedup_against_to(to_: str, cc_: str) -> str:
    to_set = {x.lower() for x in _split_recipients(to_)}
    cc_list = [x for x in _split_recipients(cc_) if x.lower() not in to_set]
    return "; ".join(cc_list)

def strip_angle_display(s: str) -> str:
    """
    If value is like "John Doe <john@acme.com>", return "john@acme.com".
    Otherwise return original.
    """
    s = str(s or "").strip()
    if "<" in s and ">" in s:
        inside = s.split("<", 1)[1].split(">", 1)[0].strip()
        return inside or s
    return s

def derive_display_name_from_email(email: str) -> str:
    email = strip_angle_display(email)
    local = str(email or "").split("@", 1)[0]
    pretty = local.replace(".", " ").replace("_", " ").replace("-", " ").strip()
    return " ".join(w.capitalize() for w in pretty.split()) or "POC"

EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")

def looks_like_email(s: str) -> bool:
    s = strip_angle_display(s)
    return bool(EMAIL_RE.match(s))

def assert_recipients_or_warn(row_idx: int, label: str, recips: str) -> bool:
    bad = []
    for r in _split_recipients(recips):
        if r == "//":
            continue  # allowed literal marker
        core = strip_angle_display(r)
        # Permit "Display Name <email@...>" OR plain email
        if "<" in r and ">" in r:
            if not looks_like_email(core):
                bad.append(r)
        else:
            if not looks_like_email(core):
                bad.append(r)
    if bad:
        st.warning(f"Row {row_idx+1}: Invalid {label} -> {', '.join(bad)}")
        return False
    return True

# -------------------------------------
# Outlook session (reused for performance)
# -------------------------------------
class OutlookSession:
    def __enter__(self):
        pythoncom.CoInitialize()
        self.app = win32com.client.Dispatch("Outlook.Application")
        return self.app
    def __exit__(self, exc_type, exc, tb):
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

def create_oft_bytes_reuse(outlook, subject: str, to_: str, cc_: str, bcc_: str, html_body: str) -> bytes:
    """
    Create .oft via Outlook COM reuse. Ensures:
    - BodyFormat = 2 (olFormatHTML)
    - SaveAs(..., 2)  -> 2 = olTemplate
    """
    mail = outlook.CreateItem(0)  # 0 = olMailItem
    mail.To = to_
    mail.CC = cc_
    mail.BCC = bcc_
    mail.Subject = subject or " "     # some Outlook versions require non-empty subject
    mail.BodyFormat = 2               # 2 = olFormatHTML
    mail.HTMLBody = html_body
    with tempfile.TemporaryDirectory() as tmp:
        path = os.path.join(tmp, "tmp.oft")
        mail.SaveAs(path, 2)          # 2 = olTemplate (.oft)
        with open(path, "rb") as f:
            return f.read()

# -------------------------------------
# Subjects & Bodies
# -------------------------------------
# ER&D — firm-approved wording (EXACT)
SUBJECT_ERD = "ER&D Data Collection - {{ case_code }} ({{ client_name }})"

BODY_SEBASTIAN_ERD = """
<p>Hi {{ case_manager_name }},</p>

<p>Hope you are doing well!</p>

<p>
I am the practice manager for Engineering and R&amp;D and I wanted to reach out regarding your work with <strong>{{ client_name }}</strong> (<strong>{{ case_code }}</strong>). From what we heard your case also included an ER&amp;D component and we would like to get your support with PI practice’s efforts in building proprietary ER&amp;D benchmarking databases.
</p>

<p>
The benchmarking team (in cc) will be reaching out with specifics. The team can help address any queries and will work with you to gather data for our Benchmarking database. If you feel that you do not have visibility for the asked information or access to client data on ER&amp;D, please let us know.
For your reference, in case there are any concerns regarding sharing sensitive client data or confidentiality, we have worked extensively with Legal, and the standard Bain MSA includes language that allows us to collect and store data for benchmarking purposes. Moreover, our Benchmarking CoE team follows a very rigorous “double blind” process that disguises and protects any client data collected. BCoE also has a “do not contact” list that tells us explicitly which clients we should not collect data from, per their contracts.
</p>

<p>
Additionally, we would also like to highlight potential benchmarking resources, please refer to the Guide to ER&amp;D Benchmarking Sources, for more details on the R&amp;D benchmarks available with our Benchmarking CoE team. We also have a wide array of benchmarks across functions like Support functions, Supply Chain and ZBB from proprietary databases (curated by Bain experts) and other third party vendors (APQC, Gartner, IFMA, ALM, Stella, MPI etc.) available with us.
</p>

<p>Thanks in advance!</p>

<p>Best,<br/>Sebastian</p>
"""

BODY_POC_ERD = """
<p>Hi {{ case_manager_name }},</p>

<p>Hope you're doing well!</p>

<p>
I work with the Benchmarking team and following up on e-mail below, we would need your support in completing the <a href="https://benchmarkingsurvey.bain.com/">linked survey</a> based on the ER&amp;D work you are doing with <strong>{{ client_name }}</strong>(<strong>{{ case_code }}</strong>). To kick-start this data collection, we have two asks from you at this point:
</p>

<ul>
<li>Identify a case team member for this task who can work with us in filling the linked survey, and we’ll provide the access link from our end</li>
<li>Set up a brief call to align on what kind of data would be available and how we can best work together on this. I can directly run through your calendar or work with your EA and find a convenient slot. Let me know what works best for you</li>
</ul>

<p>Thank you,<br/>{{ poc_display_name }}</p>

<p><em>More details on the survey</em></p>

<p><strong>Content:</strong> A high level view of the survey: You'll find instructions on the first tab and definitions throughout the survey as you click to enter data. We are collecting data across the following sections:</p>
<ul>
<li>‘Demographics' tab: Descriptors of the company or business unit in scope for Bain case – this spans basic demographics, financials, ownership, organization, and strategic/competitive position</li>
<li>‘Overall ER&amp;D Survey’ tab: Data on overall R&amp;D cost, organization layers, time spent, and performance</li>
<li>‘ER&amp;D SW Survey’ tab: More focused on software-specific metrics such as developer time, code, pull requests, development efficiency, and more. Please feel free to skip this tab if it's not relevant.</li>
</ul>

<p>Important points to note are that we are aiming to get following separate sets of data:</p>
<ul>
<li>‘As-Is' data: Client data at the start of the Bain work (would also include any estimates that the Bain case team has made which reflect the As-Is state of the client, and can be used for Benchmarking purposes)</li>
<li>‘To-Be’ data: Committed targets/recommendations, ideally the values which have been agreed to by the client based on Bain work</li>
</ul>
"""

BODY_ASEEM_ERD = """
<p>Hi {{ case_manager_name }},</p>

<p>Hope you're doing well.</p>

<p>
I lead the ER&amp;D benchmarking team at BCN and following up on the below, it would be great if you could connect us to a team member who can help us in filling the attached ER&amp;D data survey for <strong>{{ client_name }}</strong>.
</p>

<p>
If in case you’re tied up with case work, please feel free to let us know if we should get back at a later date.
</p>

<p>Looking forward to hearing from you.</p>

<p>Best,<br/>Aseem</p>
"""

# Placeholders for non-ER&D (same structure, lorem ipsum)
def _lipsum_initial(func_label: str) -> str:
    return f"""
<p>Hi {{ {{ case_manager_name }} }},</p>

<p>Hope you are doing well!</p>

<p>
I am writing regarding your work with <strong>{{{{ client_name }}}}</strong> (<strong>{{{{ case_code }}}}</strong>) and our ongoing Data Collection initiative for <strong>{func_label}</strong>. Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.
</p>

<p>
The benchmarking team (in cc) will follow up with specifics and support throughout the process. In case of any concerns about data handling or confidentiality, please note we follow a rigorous process to protect client information. Lorem ipsum dolor sit amet, consectetur adipiscing elit.
</p>

<p>Thanks in advance!</p>

<p>Best,<br/>Sebastian</p>
""".strip()

def _lipsum_poc(func_label: str) -> str:
    return f"""
<p>Hi {{ {{ case_manager_name }} }},</p>

<p>Hope you're doing well!</p>

<p>
Following up on the note below, we would appreciate your support in completing the <a href="#">linked survey</a> for <strong>{func_label}</strong> based on the work with <strong>{{{{ client_name }}}}</strong> (<strong>{{{{ case_code }}}}</strong>). To kick-start, we have two quick asks:
</p>

<ul>
<li>Identify a team member who can work with us to fill the survey; we will share access from our end.</li>
<li>Set up a brief call to align on available data and the best way to collaborate. Lorem ipsum dolor sit amet.</li>
</ul>

<p>Thank you,<br/>{{{{ poc_display_name }}}}</p>

<p><em>More details on the survey</em></p>

<p><strong>Content:</strong> Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sections include basic demographics, process measures, and performance indicators relevant to {func_label}.</p>
<ul>
<li>‘Demographics’ tab</li>
<li>‘Overall {func_label} Survey’ tab</li>
<li>‘{func_label} Advanced’ tab (optional)</li>
</ul>

<p>We aim to collect both ‘As-Is’ and ‘To-Be’ data. Lorem ipsum dolor sit amet.</p>
""".strip()

def _lipsum_escalation(func_label: str) -> str:
    return f"""
<p>Hi {{ {{ case_manager_name }} }},</p>

<p>Hope you're doing well.</p>

<p>
Following up on the below, it would be great if you could connect us to a team member who can help fill the attached data survey for <strong>{{{{ client_name }}}}</strong> ({func_label}). Lorem ipsum dolor sit amet, consectetur adipiscing elit.
</p>

<p>
If you're tied up with case work, happy to reconnect at a later date. Looking forward to hearing from you.
</p>

<p>Best,<br/>Aseem</p>
""".strip()

# Function set
FUNCTIONS = ["ER&D", "Supply Chain", "Procurement", "Manufacturing"]

def get_templates_for_function(func_name: str) -> Dict[str, str]:
    """Return subject + bodies for the selected function.
       For ER&D: exact wording; others: structure-matched lorem ipsum placeholders."""
    if func_name == "ER&D":
        return {
            "subject": SUBJECT_ERD,
            "sebastian": BODY_SEBASTIAN_ERD,
            "poc": BODY_POC_ERD,
            "aseem": BODY_ASEEM_ERD,
            "subject_label": "ER&D",
        }
    # Non-ER&D placeholders with function-specific subject
    subject = f"{func_name} Data Collection - {{ {{ case_code }} }} ({{ {{ client_name }} }})"
    return {
        "subject": subject,
        "sebastian": _lipsum_initial(func_name),
        "poc": _lipsum_poc(func_name),
        "aseem": _lipsum_escalation(func_name),
        "subject_label": func_name,
    }

# -------------------------------------
# Generic DC generator (templated per function)
# -------------------------------------
def run_oft_generator_for_function(func_name: str):
    tpls = get_templates_for_function(func_name)
    SUBJECT_COMMON = tpls["subject"]
    BODY_SEBASTIAN = tpls["sebastian"]
    BODY_POC = tpls["poc"]
    BODY_ASEEM = tpls["aseem"]
    subject_label = tpls["subject_label"]

    st.caption(f"Outputs real .oft files only (no EML). Subjects are identical for all three; BCC is '//' for all. Function: **{subject_label}**")

    with st.expander("Required Excel columns & rules", expanded=False):
        st.markdown(
            f"""
**Required columns:** `client_name`, `case_code`, `case_manager_name`, `to`, `team_lead_email`, `POC_name`  
**Optional:** `POC_display_name`, `extra_cc`

**To / CC / BCC rules per template** *(same as ER&D for now; customize later as needed)*:
- **Sebastian (Initial)** — To: `to`; CC: `team_lead_email` + `POC_name` + `extra_cc` + `ERDDBTeam.Global@Bain.com`; **BCC:** `//`
- **POC (Follow-up)** — To: `to`; CC: `ERDDBTeam.Global@Bain.com` + `team_lead_email` + `Sebastian.Sambale@Bain.com`; **BCC:** `//`
- **Aseem (Escalation)** — To: `to`; CC: `Sebastian.Sambale@Bain.com` + `team_lead_email` + `ERDDBTeam.Global@Bain.com` + `POC_name`; **BCC:** `//`

**Subject (same pattern for all):**  
`{subject_label} Data Collection - {{ {{ case_code }} }} ({{ {{ client_name }} }})`
            """
        )

    # Excel upload
    st.subheader("1) Upload Excel (.xlsx)")
    excel_file = st.file_uploader("Upload a .xlsx file (first sheet will be used)", type=["xlsx"], key=f"uploader_{func_name}")

    rows: List[dict] = []
    if excel_file is not None:
        try:
            df = pd.read_excel(excel_file, dtype=str, keep_default_na=False).fillna("")
            required = ["client_name", "case_code", "case_manager_name", "to", "team_lead_email", "POC_name"]
            missing = [c for c in required if c not in df.columns]
            if missing:
                st.error(f"Excel is missing columns: {', '.join(missing)}")
                st.stop()
            if "POC_display_name" not in df.columns:
                df["POC_display_name"] = ""
            if "extra_cc" not in df.columns:
                df["extra_cc"] = ""
            st.dataframe(df[required + ["POC_display_name", "extra_cc"]], use_container_width=True)
            rows = df.to_dict(orient="records")
        except Exception as e:
            st.error(f"Could not read Excel: {e}")

    # Preview
    if rows:
        st.subheader("2) Preview (from first row)")
        r0 = rows[0]
        # Safe strings
        client_name      = str(r0.get("client_name") or "").strip()
        case_code        = str(r0.get("case_code") or "").strip()
        case_manager     = str(r0.get("case_manager_name") or "").strip()
        to_cm            = str(r0.get("to") or "").strip()
        team_lead_email  = str(r0.get("team_lead_email") or "").strip()
        poc_email        = str(r0.get("POC_name") or "").strip()
        poc_display_name = (str(r0.get("POC_display_name") or "").strip()
                            or derive_display_name_from_email(poc_email))
        extra_cc         = str(r0.get("extra_cc") or "").strip()
        bcc_preview      = "//"

        ctx = {
            "client_name": client_name,
            "case_code": case_code,
            "case_manager_name": case_manager,
            "poc_display_name": poc_display_name,
            "today": datetime.now().strftime("%d %b %Y"),
        }
        subject_preview = render_text(SUBJECT_COMMON, ctx)

        # --- Sebastian ---
        cc_seb = build_cc(team_lead_email, poc_email, extra_cc, "ERDDBTeam.Global@Bain.com")
        cc_seb = dedup_against_to(to_cm, cc_seb)
        st.markdown("#### Sebastian – Initial")
        st.write(f"**To:** {to_cm}")
        st.write(f"**CC:** {cc_seb}")
        st.write(f"**BCC:** {bcc_preview}")
        st.write(f"**Subject:** {subject_preview}")
        st.markdown(render_text(BODY_SEBASTIAN, ctx), unsafe_allow_html=True)
        st.divider()

        # --- POC ---
        cc_poc = build_cc("ERDDBTeam.Global@Bain.com", team_lead_email, "Sebastian.Sambale@Bain.com")
        cc_poc = dedup_against_to(to_cm, cc_poc)
        st.markdown("#### POC – Follow-up")
        st.write(f"**To:** {to_cm}")
        st.write(f"**CC:** {cc_poc}")
        st.write(f"**BCC:** {bcc_preview}")
        st.write(f"**Subject:** {subject_preview}")
        st.markdown(render_text(BODY_POC, ctx), unsafe_allow_html=True)
        st.divider()

        # --- Aseem ---
        cc_aseem = build_cc("Sebastian.Sambale@Bain.com", team_lead_email, "ERDDBTeam.Global@Bain.com", poc_email)
        cc_aseem = dedup_against_to(to_cm, cc_aseem)
        st.markdown("#### Aseem – Escalation")
        st.write(f"**To:** {to_cm}")
        st.write(f"**CC:** {cc_aseem}")
        st.write(f"**BCC:** {bcc_preview}")
        st.write(f"**Subject:** {subject_preview}")
        st.markdown(render_text(BODY_ASEEM, ctx), unsafe_allow_html=True)

        st.dataframe(pd.DataFrame([{
            "Template": "Sebastian – Initial", "To": to_cm, "CC": cc_seb, "BCC": bcc_preview
        },{
            "Template": "POC – Follow-up", "To": to_cm, "CC": cc_poc, "BCC": bcc_preview
        },{
            "Template": "Aseem – Escalation", "To": to_cm, "CC": cc_aseem, "BCC": bcc_preview
        }]), use_container_width=True)

    # Generate
    st.subheader("3) Generate & Download (real OFT)")
    if st.button("Generate .oft templates", key=f"btn_gen_{func_name}"):
        if not rows:
            st.warning("Please upload Excel first.")
        elif not WINDOWS or not HAS_WIN32:
            st.error("This app requires Windows + Outlook (pywin32). Please run on a Windows machine with Outlook installed.")
        else:
            mem = io.BytesIO()
            status = []
            with zipfile.ZipFile(mem, "w", zipfile.ZIP_DEFLATED) as zf, OutlookSession() as outlook:
                for i, r in enumerate(rows):
                    try:
                        client = str(r.get("client_name") or "").strip()
                        code = str(r.get("case_code") or "").strip()
                        cm = str(r.get("case_manager_name") or "").strip()
                        to_ = str(r.get("to") or "").strip()
                        tl = str(r.get("team_lead_email") or "").strip()
                        poc = str(r.get("POC_name") or "").strip()
                        poc_disp = (str(r.get("POC_display_name") or "").strip()
                                    or derive_display_name_from_email(poc))
                        extra = str(r.get("extra_cc") or "").strip()

                        if not (client and code and to_):
                            status.append(f"Row {i+1}: SKIPPED – missing client/code/to")
                            continue

                        ctx = {
                            "client_name": client,
                            "case_code": code,
                            "case_manager_name": cm,
                            "poc_display_name": poc_disp,
                            "today": datetime.now().strftime("%d %b %Y"),
                        }
                        subject = render_text(SUBJECT_COMMON, ctx)
                        base = sanitize_filename(f"{client} - {code}")
                        bcc_val = "//"

                        # Build & dedup recipients (rules same as ER&D for now)
                        cc1_build = build_cc(tl, poc, extra, "ERDDBTeam.Global@Bain.com")
                        cc2_build = build_cc("ERDDBTeam.Global@Bain.com", tl, "Sebastian.Sambale@Bain.com")
                        cc3_build = build_cc("Sebastian.Sambale@Bain.com", tl, "ERDDBTeam.Global@Bain.com", poc)

                        cc1 = dedup_against_to(to_, cc1_build)
                        cc2 = dedup_against_to(to_, cc2_build)
                        cc3 = dedup_against_to(to_, cc3_build)

                        ok = True
                        ok &= assert_recipients_or_warn(i, "To", to_)
                        ok &= assert_recipients_or_warn(i, "CC (Sebastian)", cc1)
                        ok &= assert_recipients_or_warn(i, "CC (POC)", cc2)
                        ok &= assert_recipients_or_warn(i, "CC (Aseem)", cc3)
                        if not ok:
                            status.append(f"Row {i+1}: SKIPPED – invalid recipients")
                            continue

                        # 1) Sebastian
                        body = render_text(BODY_SEBASTIAN, ctx)
                        oft_bytes = create_oft_bytes_reuse(outlook, subject, to_, cc1, bcc_val, body)
                        zf.writestr(f"1_Sebastian_Initial/{base}.oft", oft_bytes)

                        # 2) POC
                        body = render_text(BODY_POC, ctx)
                        oft_bytes = create_oft_bytes_reuse(outlook, subject, to_, cc2, bcc_val, body)
                        zf.writestr(f"2_POC_Follow_Up/{base}.oft", oft_bytes)

                        # 3) Aseem
                        body = render_text(BODY_ASEEM, ctx)
                        oft_bytes = create_oft_bytes_reuse(outlook, subject, to_, cc3, bcc_val, body)
                        zf.writestr(f"3_Aseem_Escalation/{base}.oft", oft_bytes)

                        status.append(f"Row {i+1}: OK – {base}")
                    except Exception as e:
                        status.append(f"Row {i+1}: FAILED – {e}")

            st.download_button(
                "⬇️ Download ZIP (three OFT folders)",
                data=mem.getvalue(),
                file_name=f"bain-{func_name.lower().replace(' ', '')}-ofts_{datetime.now():%Y%m%d_%H%M%S}.zip",
                mime="application/zip",
            )
            st.success(f"Generated real .oft files with the exact bodies (placeholders for {func_name} unless ER&D), same subject pattern, and BCC='//'. Rows processed: {len(rows)}")
            st.text("\n".join(status[:500]))

# -------------------------------------
# App Shell: Function selection
# -------------------------------------
st.set_page_config(page_title="Multi-Function DC OFT Generator", page_icon="📧", layout="centered")

st.title("📧 Data Collection OFT Generator (Multi-Function)")
st.write("Select a function to begin. ER&D uses approved copy; other functions use structure-matched placeholders for now.")

# Sidebar selection
with st.sidebar:
    st.header("Choose Function")
    chosen = st.radio("Function", FUNCTIONS, index=0)

# Guard: environment
if not WINDOWS or not HAS_WIN32:
    st.error("This app requires Windows + Outlook (pywin32). Please run on a Windows machine with Outlook installed.")
    st.stop()

# Render the chosen function's generator
run_oft_generator_for_function(chosen)
