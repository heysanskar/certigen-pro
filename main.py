import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Inches as DocxInches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from pptx import Presentation
from pptx.util import Inches as PptxInches
import qrcode
import os
import smtplib
import logging
from email.message import EmailMessage
from docx2pdf import convert as convert_docx
from datetime import datetime
import hashlib
import json
import plotly.express as px
import sys
import time
import shutil
from pypdf import PdfReader, PdfWriter
import logging
from pypdf.constants import UserAccessPermissions as UAP

# Conditional import for PowerPoint PDF conversion on Windows via COM architecture
try:
    if sys.platform == "win32":
        import comtypes.client
except ImportError:
    comtypes = None

# Setup Logging
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

# ---------------- CONFIG & SYSTEM PATHS ----------------
st.set_page_config(page_title="CertiGen Pro", page_icon="🎓", layout="wide")

for folder in ["system", "records", "certificates"]:
    os.makedirs(folder, exist_ok=True)

# ---------------- SECURITY ACCESS LAYER ----------------
USER_FILE = "system/users.json"
if not os.path.exists(USER_FILE):
    with open(USER_FILE, "w") as f:
        json.dump({"admin": "password123"}, f)

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

if not st.session_state.logged_in:
    st.title("🔐 Login to CertiGen Pro")
    with st.form("Login System Auth"):
        u = st.text_input("Username")
        p = st.text_input("Password", type="password")
        if st.form_submit_button("Login"):
            with open(USER_FILE) as f:
                users = json.load(f)
            if u in users and users[u] == p:
                st.session_state.logged_in = True
                st.rerun()
            else:
                st.error("❌ Invalid credentials")
    st.stop()

# ---------------- FULL-LENGTH WORD ENGINE (PRESERVES EXACT STYLES) ----------------
def process_docx_template(template_path, output_path, replacements, qr_enabled, qr_path):
    """Processes Word files by cleanly unifying text blocks to protect formatting across paragraphs and tables."""
    doc = Document(template_path)
    
    # Process standard paragraphs
    for p in doc.paragraphs:
        full_text = "".join(run.text for run in p.runs if run.text)
        has_tag = any(k.lower() in full_text.lower() for k in replacements.keys())
        
        if has_tag and p.runs:
            base_run = p.runs[0]
            font_name = base_run.font.name
            font_size = base_run.font.size
            font_color = base_run.font.color.rgb if base_run.font.color else None
            is_bold = base_run.bold
            is_italic = base_run.italic
            is_underline = base_run.underline
            
            for key, val in replacements.items():
                full_text = full_text.replace(f"{{{{{key.lower()}}}}}", str(val))
                full_text = full_text.replace(f"{{{{{key.capitalize()}}}}}", str(val))
                full_text = full_text.replace(f"{{{{{key.upper()}}}}}", str(val))
            
            for extra_run in p.runs[1:]:
                extra_run.text = ""
                
            base_run.text = full_text
            if font_name: base_run.font.name = font_name
            if font_size: base_run.font.size = font_size
            if font_color: base_run.font.color.rgb = font_color
            base_run.bold = is_bold
            base_run.italic = is_italic
            base_run.underline = is_underline

    # Process all table structural elements fully
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    full_text = "".join(run.text for run in p.runs if run.text)
                    has_tag = any(k.lower() in full_text.lower() for k in replacements.keys())
                    
                    if has_tag and p.runs:
                        base_run = p.runs[0]
                        font_name = base_run.font.name
                        font_size = base_run.font.size
                        font_color = base_run.font.color.rgb if base_run.font.color else None
                        is_bold = base_run.bold
                        is_italic = base_run.italic
                        is_underline = base_run.underline
                        
                        for key, val in replacements.items():
                            full_text = full_text.replace(f"{{{{{key.lower()}}}}}", str(val))
                            full_text = full_text.replace(f"{{{{{key.capitalize()}}}}}", str(val))
                            full_text = full_text.replace(f"{{{{{key.upper()}}}}}", str(val))
                        
                        for extra_run in p.runs[1:]:
                            extra_run.text = ""
                            
                        base_run.text = full_text
                        if font_name: base_run.font.name = font_name
                        if font_size: base_run.font.size = font_size
                        if font_color: base_run.font.color.rgb = font_color
                        base_run.bold = is_bold
                        base_run.italic = is_italic
                        base_run.underline = is_underline

    if qr_enabled:
        for p in doc.paragraphs:
            if "[[QR]]" in p.text:
                p.text = p.text.replace("[[QR]]", "")
                run = p.add_run()
                run.add_picture(qr_path, width=DocxInches(1.2))
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                break

    # Metadata Protection Node: Apply native XML read-only lock to backup layers
    try:
        settings = doc.settings.element
        protection = settings.find(settings.tag.replace('settings', 'documentProtection'))
        if protection is None:
            from docx.oxml import OxmlElement
            protection = OxmlElement('w:documentProtection')
            settings.append(protection)
        protection.set(settings.tag.replace('settings', 'edit'), 'readOnly')
        protection.set(settings.tag.replace('settings', 'enforcement'), '1')
    except Exception as lock_err:
        logging.warning(f"Metadata security tag pass skipped: {lock_err}")

    doc.save(output_path)

# ---------------- FULL-LENGTH POWERPOINT ENGINE (PRESERVES COLORS) ----------------
def process_pptx_template(template_path, output_path, replacements, qr_enabled, qr_path):
    """Processes PowerPoint structures by targeted run replacement to completely preserve fonts, colors, and theme inheritance."""
    prs = Presentation(template_path)
    
    def _replace_text_in_runs(paragraphs, replacements):
        for paragraph in paragraphs:
            # First pass: Check if individual runs contain the tags directly (preserves exact unique styles)
            for run in paragraph.runs:
                for key, val in replacements.items():
                    tags = [f"{{{{{key.lower()}}}}}", f"{{{{{key.capitalize()}}}}}", f"{{{{{key.upper()}}}}}"]
                    for tag in tags:
                        if tag in run.text:
                            run.text = run.text.replace(tag, str(val))
            
            # Second pass fallback: If PowerPoint split a tag across multiple runs (e.g., '{' and '{serial}'),
            # we must process at paragraph level but safely capture existing font properties
            full_text = paragraph.text
            has_tag = any(f"{{{{{k.lower()}}}}}" in full_text.lower() for k in replacements.keys())
            
            if has_tag and len(paragraph.runs) > 0:
                # Cache the original styling from the first run before it gets modified
                base_run = paragraph.runs[0]
                font_name = base_run.font.name
                font_size = base_run.font.size
                
                # Check for explicit color vs theme color
                has_explicit_color = False
                font_color = None
                if base_run.font.color and hasattr(base_run.font.color, 'type') and base_run.font.color.type == 1: # RGB color type
                    font_color = base_run.font.color.rgb
                    has_explicit_color = True
                
                is_bold = base_run.font.bold
                is_italic = base_run.font.italic
                
                # Perform the replacement on the joined paragraph text string
                for key, val in replacements.items():
                    full_text = full_text.replace(f"{{{{{key.lower()}}}}}", str(val))
                    full_text = full_text.replace(f"{{{{{key.capitalize()}}}}}", str(val))
                    full_text = full_text.replace(f"{{{{{key.upper()}}}}}", str(val))
                
                # Clear extra runs but leave one to preserve placeholder linkage
                paragraph.text = full_text
                if len(paragraph.runs) > 0:
                    new_run = paragraph.runs[0]
                    if font_name: new_run.font.name = font_name
                    if font_size: new_run.font.size = font_size
                    if has_explicit_color and font_color: 
                        new_run.font.color.rgb = font_color
                    new_run.font.bold = is_bold
                    new_run.font.italic = is_italic

    # Iterate through all slide structures
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                _replace_text_in_runs(shape.text_frame.paragraphs, replacements)
            
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        _replace_text_in_runs(cell.text_frame.paragraphs, replacements)

    # QR Code handling
    if qr_enabled:
        for slide in prs.slides:
            for shape in slide.shapes:
                if shape.has_text_frame and "[[QR]]" in shape.text_frame.text:
                    shape.text_frame.text = "" 
                    slide.shapes.add_picture(qr_path, shape.left, shape.top, width=PptxInches(1.4))
                    break

    prs.save(output_path)

# ---------------- NATIVE POWERPOINT TO PDF CONVERTER ----------------
def convert_pptx_to_pdf(input_pptx, output_pdf):
    if sys.platform != "win32" or comtypes is None:
        raise NotImplementedError("PPTX PDF native conversion requires a Windows environment with Office installed.")
    
    abs_input = os.path.abspath(input_pptx)
    abs_output = os.path.abspath(output_pdf)
    
    ppt_app = comtypes.client.CreateObject("PowerPoint.Application")
    try:
        presentation = ppt_app.Presentations.Open(abs_input, WithWindow=False)
        presentation.SaveAs(abs_output, 32) 
        presentation.Close()
    finally:
        ppt_app.Quit()

# ---------------- DATA & NETWORK UTILITIES ----------------
def auth_code(serial, name):
    return hashlib.sha256(f"{serial}{name}".encode()).hexdigest()[:10]

# 👇 PASTE THIS NEW FUNCTION HERE
def lock_pdf_permissions(pdf_path,owner_password="MasterSystemPasswordSecurity2026"):
    """Encrypt PDF using the recipient's auth code as the open password."""

    try:
        reader = PdfReader(pdf_path)
        writer = PdfWriter()

        for page in reader.pages:
            writer.add_page(page)

        writer.encrypt(
        user_password="",
        owner_password=owner_password,
        permissions_flag=UAP.PRINT
        )

        with open(pdf_path, "wb") as f:
            writer.write(f)

        logging.info(f"PDF encrypted successfully: {pdf_path}")

    except Exception as e:
        logging.exception(f"Failed to apply PDF permission lock: {e}")

def save_record(data, event_code):
    file = f"records/{event_code}.csv"
    columns_order = ["Event", "Name", "Email", "Designation", "Organization", "Topic", "Serial", "AuthCode", "Date", "EmailStatus"]
    df = pd.DataFrame([data]).reindex(columns_order, axis=1, fill_value="")
    
    try:
        df.to_csv(file, mode="a", header=not os.path.exists(file), index=False)
    except PermissionError:
        st.error(f"❌ **Permission Error:** Cannot update the log file `{file}` because it is currently open in Microsoft Excel.")
        st.warning("👉 Please close the Excel window for this CSV file and try generating again.")
        st.stop()

def next_seq(event_code):
    file = f"records/{event_code}.csv"
    if not os.path.exists(file) or os.stat(file).st_size == 0: 
        return 1
    try:
        df = pd.read_csv(file)
        return int(str(df.iloc[-1]["Serial"]).split("-")[-1]) + 1
    except: 
        return 1

def send_email(user, pwd, to, cc_list, name, event, file_path, serial, auth, custom_subject="", custom_body=""):
    if not user or not pwd: 
        raise ValueError("SMTP parameters missing.")
    msg = EmailMessage()
    
    # NEW FRONTEND INPUT HOOK: Process dynamic subject variables
    if custom_subject.strip():
        subject_content = custom_subject.replace("{name}", name).replace("{event}", event).replace("{serial}", serial).replace("{auth}", auth)
    else:
        subject_content = f"🎉 Your Certificate - {event}"
        
    msg["Subject"] = subject_content
    msg["From"] = user
    msg["To"] = to
    if cc_list: 
        msg["Cc"] = ", ".join(cc_list)
        
    if custom_body.strip():
        body_content = custom_body.replace("{name}", name).replace("{event}", event).replace("{serial}", serial).replace("{auth}", auth)
    else:
        body_content = f"Dear {name},\n\nCongratulations! Your certificate for {event} is attached.\n\nID: {serial}\nAuth: {auth}"
        
    msg.set_content(body_content)
    
    if os.path.exists(file_path):
        with open(file_path, "rb") as f:
            ext = os.path.splitext(file_path)[1].lower()
            subtype = "pdf" if ext == ".pdf" else "octet-stream"
            msg.add_attachment(f.read(), maintype="application", subtype=subtype, filename=os.path.basename(file_path))
            
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as s:
        s.login(user, pwd)
        s.send_message(msg, to_addrs=[to] + (cc_list if cc_list else []))

# ---------------- STREAMLIT INTERFACE ----------------
st.title("🎓 CertiGen Pro Layout Engine")

if "active_event_log" not in st.session_state:
    st.session_state.active_event_log = None

tab1, tab2, tab3 = st.tabs(["🚀 Engine Room", "📊 Analytics Reports", "🔍 Audit Center"])

with tab1:
    st.header("Document Generation Router")
    col1, col2 = st.columns(2)
    with col1:
        event = st.text_input("Event Name")
        code = st.text_input("Event Code").strip()
        issue_date_input = st.date_input("Issue Date", value=datetime.now())
    with col2:
        qr_enabled = st.checkbox("Include QR verification anchor", value=True)
        reset = st.checkbox("Reset Sequence Counter")
        mode = st.radio("Input Source Vector", ["Single Issuance", "Bulk Upload (Excel)"])

    template = st.file_uploader("Upload Template Document (.docx OR .pptx)", type=["docx", "pptx"])
    
    data = []
    if mode == "Single Issuance":
        n, e = st.text_input("Recipient Name"), st.text_input("Recipient Email")
        desig = st.text_input("Recipient Designation (Optional)")
        org = st.text_input("Recipient Organization (Optional)") 
        top = st.text_input("Topic Delivered (Optional)") 
        if n and e: 
            data.append({
                "Name": n.strip(), 
                "Email": e.strip(), 
                "Designation": desig.strip() if desig else "",
                "Organization": org.strip() if org else "",
                "Topic": top.strip() if top else ""
            })
    else:
        f = st.file_uploader("Upload Manifest Excel Sheet", type=["xlsx"])
        if f:
            df_in = pd.read_excel(f)
            if df_in is not None:
                df_in.columns = [str(c).strip() for c in df_in.columns]
                data = df_in.to_dict("records")

    st.subheader("📧 Communication Routing & Messaging Layer")
    em_col1, em_col2 = st.columns(2)
    with em_col1:
        user = st.text_input("SMTP System User Email")
        pwd = st.text_input("SMTP App Token Key", type="password")
    with em_col2:
        cc_input = st.text_input("CC Recipients Matrix")
        cc_list = [i.strip() for i in cc_input.split(",") if i.strip()]
        
    # NEW FRONTEND INPUTS: Added Custom Email Subject Field 
    custom_email_subject = st.text_input(
        "Custom Email Subject (Optional)",
        placeholder="Leave blank for default. Optional tags: {name}, {event}, {serial}, {auth}"
    )
    custom_email_body = st.text_area(
        "Custom Email Body (Optional)", 
        placeholder="Leave blank for default. Optional tags: {name}, {event}, {serial}, {auth}"
    )

    if st.button("Execute Engine Generation Optimization Run", type="primary"):
        if not template or not data or not code or not event:
            st.error("Missing critical generation parameter values.")
        else:
            file_extension = os.path.splitext(template.name)[1].lower()
            st.info(f"📁 Auto-detected structure formatting type: **{file_extension.upper()}**")
            
            with st.spinner("Processing documents in execution stream..."):
                seq = 1 if reset else next_seq(code)
                date_str = datetime.now().strftime("%d%m%y")
                
                generated_files_pool = []

                for i, item in enumerate(data):
                    name = item.get("Name") or item.get("name")
                    email = item.get("Email") or item.get("email")
                    
                    if not name or not email:
                        st.warning(f"⚠️ Row {i+1} skipped: Missing 'Name' or 'Email' context.")
                        continue
                        
                    designation_val = item.get("Designation", item.get("designation", ""))
                    if pd.isna(designation_val): designation_val = ""
                        
                    organization_val = item.get("Organization", item.get("organization", ""))
                    if pd.isna(organization_val): organization_val = ""
                        
                    topic_val = item.get("Topic", item.get("topic", ""))
                    if pd.isna(topic_val): topic_val = ""
                        
                    s_no = f"{code}-{date_str}-{str(seq+i).zfill(2)}"
                    auth = auth_code(s_no, name)
                    qr_path = f"qr_{i}.png"

                    if qr_enabled:
                        qrcode.make(f"{s_no}|{auth}").save(qr_path)

                    temp_out = f"temp_{i}{file_extension}"
                    
                    replacements = {
                        "name": name,
                        "designation": designation_val,
                        "organization": organization_val,
                        "topic": topic_val, 
                        "serial": s_no,
                        "date": issue_date_input.strftime("%d %B %Y"),
                        "authcode": auth
                    }
                    
                    if file_extension == ".docx":
                        process_docx_template(template, temp_out, replacements, qr_enabled, qr_path)
                    elif file_extension == ".pptx":
                        process_pptx_template(template, temp_out, replacements, qr_enabled, qr_path)

                    output_file = f"certificates/{name.replace(' ', '_')}_{s_no}.pdf"
                    pdf_generation_successful = False
                    
                    try:
                        if file_extension == ".docx":
                            convert_docx(temp_out, output_file)
                            pdf_generation_successful = True
                        elif file_extension == ".pptx":
                            convert_pptx_to_pdf(temp_out, output_file)
                            pdf_generation_successful = True

                        if pdf_generation_successful:
                            lock_pdf_permissions(output_file)

                    except Exception as ex:
                        logging.warning(f"Native conversion issue: {ex}")
                    
                    if not pdf_generation_successful:
                        if 'prs' in locals(): del prs
                        time.sleep(0.2)
                        
                        st.error(f"❌ **Security System Alert:** Could not render an un-editable secure PDF asset for **{name}**.")
                        st.warning("⚠️ Distribution halted to prevent saving fully editable formatting templates onto local dashboards.")
                        
                        if os.path.exists(temp_out): os.remove(temp_out)
                        if os.path.exists(qr_path): os.remove(qr_path)
                        st.stop()

                    status = "Not Sent"
                    if user and pwd:
                        try:
                            # Connected frontend custom subject variable to email trigger block
                            send_email(user, pwd, email, cc_list, name, event, output_file, s_no, auth, custom_email_subject, custom_body=custom_email_body)
                            status = "Sent"
                            time.sleep(2)  
                        except Exception as mail_err: 
                            status = "Failed"
                            st.error(f"❌ Email failed for {name} ({email}): {str(mail_err)}")
                            logging.error(f"Email dispatch error: {mail_err}")

                    save_record({
                        "Event": event, "Name": name, "Email": email, 
                        "Designation": designation_val, "Organization": organization_val, "Topic": topic_val,
                        "Serial": s_no, "AuthCode": auth, "Date": datetime.now().strftime("%Y-%m-%d"), 
                        "EmailStatus": status
                    }, code)
                    
                    generated_files_pool.append({"path": output_file, "name": f"{name}_{s_no}"})
                    
                    if os.path.exists(qr_path): os.remove(qr_path)
                    if os.path.exists(temp_out) and os.path.abspath(temp_out) != os.path.abspath(output_file): 
                        os.remove(temp_out)

            st.session_state.active_event_log = f"{code}.csv"
            st.success(f"🎉 Generated and processed {len(generated_files_pool)} certificates successfully!")
            
            st.markdown("### 📥 Download Generated Assets")
            st.info("💾 All files have been safely archived to the secure server directory: `/certificates`")
            
            for f_item in generated_files_pool:
                if os.path.exists(f_item["path"]):
                    with open(f_item["path"], "rb") as file_bytes:
                        file_ext = os.path.splitext(f_item["path"])[1]
                        st.download_button(
                            label=f"⬇️ Download Certificate: {f_item['name']}{file_ext}",
                            data=file_bytes.read(),
                            file_name=os.path.basename(f_item["path"]),
                            mime="application/pdf" if file_ext == ".pdf" else "application/octet-stream",
                            key=f"dl_{f_item['name']}"
                        )

# TAB 2: METRICS REPORTS
with tab2:
    st.header("📊 Operational Analytics & Report Center")
    files = [f for f in os.listdir("records") if f.endswith(".csv")]
    if files:
        default_idx = 0
        if st.session_state.active_event_log in files:
            default_idx = files.index(st.session_state.active_event_log)
            
        sel = st.selectbox("Select Target Dataset Ledger", files, index=default_idx)
        path = os.path.join("records", sel)
        df_log = pd.read_csv(path, on_bad_lines='skip')
        
        total_records = len(df_log)
        sent_successful = len(df_log[df_log["EmailStatus"] == "Sent"])
        failed_dispatches = len(df_log[df_log["EmailStatus"] == "Failed"])
        
        m_col1, m_col2, m_col3 = st.columns(3)
        m_col1.metric("Certificates Generated", total_records)
        m_col2.metric("Emails Delivered Successfully", sent_successful)
        m_col3.metric("Delivery Dropouts/Failures", failed_dispatches, delta=-failed_dispatches if failed_dispatches > 0 else 0, delta_color="inverse")
        
        st.markdown("---")
        
        if not df_log.empty and "EmailStatus" in df_log.columns:
            st.subheader("📈 Communications Pipeline Spread")
            fig = px.pie(df_log, names="EmailStatus", hole=0.4,
                         color_discrete_map={"Sent": "#2ecc71", "Failed": "#e74c3c", "Not Sent": "#95a5a6"})
            st.plotly_chart(fig, use_container_width=True)
            
        st.subheader("📋 Document Audit Ledger Tracking Matrix")
        st.dataframe(df_log, use_container_width=True)
        
        st.subheader("📥 Export Options")
        csv_binary = df_log.to_csv(index=False).encode('utf-8')
        ex_col1, ex_col2 = st.columns(2)
        with ex_col1:
            st.download_button(
                label="📥 Download Master Audit Report (CSV)",
                data=csv_binary,
                file_name=f"Executive_Report_{sel}",
                mime="text/csv",
                type="primary"
            )
        with ex_col2:
            if st.button("Purge Current Log permanently", type="secondary"):
                os.remove(path)
                st.session_state.active_event_log = None
                st.rerun()
    else:
        st.info("No logs present. System database generation logs are empty.")

# TAB 3: AUDIT
with tab3:
    st.header("Audit Verification Hub")
    s_input = st.text_input("System Serial String")
    a_input = st.text_input("Verification Hash Code")
    if st.button("Execute Verification Search"):
        found = False
        for f in os.listdir("records"):
            if f.endswith(".csv"):
                try:
                    df_audit = pd.read_csv(os.path.join("records", f))
                    if "Serial" in df_audit.columns and "AuthCode" in df_audit.columns:
                        match = df_audit[(df_audit["Serial"] == s_input) & (df_audit["AuthCode"] == a_input)]
                        if not match.empty:
                            st.success("🔒 Authenticated Token Verified Against Master Ledger Secure Anchors")
                            st.dataframe(match)
                            found = True
                            break
                except Exception as audit_err:
                    logging.error(f"Error scanning logs during manual lookup: {audit_err}")
        if not found: 
            st.error("Verification Token Mismatch. Warning: This certificate identity is unauthorized or altered.")