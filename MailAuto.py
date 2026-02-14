#import time
#import zipfile
#import pandas as pd
#import smtplib
#import streamlit as st
#from email.mime.multipart import MIMEMultipart
#from email.mime.text import MIMEText
#from email.mime.base import MIMEBase
#from email import encoders

# ==============================
# PAGE CONFIG
# ==============================
st.set_page_config(page_title="PAN Document Mailer", layout="wide")
st.title("📧 Bulk Email Sender")

# ==============================
# SIDEBAR CONFIG
# ==============================
st.sidebar.header("Email Configuration")

smtp_server = st.sidebar.text_input("SMTP Server", value="smtp.gmail.com")
smtp_port = st.sidebar.number_input("SMTP Port", value=587)

sender_email = st.sidebar.text_input("Sender Gmail")
sender_password = st.sidebar.text_input("App Password (16-digit)", type="password")

# ==============================
# EMAIL BODY TEMPLATE
# ==============================
BODY_TEMPLATE = """Dear {Name},

Please find attached the document(s) corresponding to your PAN: {PAN}.
If you have any questions, please reply to this email.

Best Regards,
Account Team
"""

# ==============================
# FILE UPLOAD SECTION
# ==============================
st.header("Upload Files")

uploaded_excel = st.file_uploader("Upload Excel File", type=["xlsx"])

st.subheader("Upload Folder Part 1 (ZIP)")
zip_1a = st.file_uploader("Upload 1A.zip", type=["zip"], key="zip1")

st.subheader("Upload Folder Part 2 (ZIP)")
zip_2a = st.file_uploader("Upload 2A.zip", type=["zip"], key="zip2")

# ==============================
# EXTRACT PDF FILES FROM ZIP
# ==============================
def extract_pdfs_from_zip(zip_file):
    extracted_files = []

    if zip_file is None:
        return extracted_files

    with zipfile.ZipFile(zip_file, 'r') as z:
        for file_name in z.namelist():
            if file_name.lower().endswith(".pdf"):
                file_data = z.read(file_name)
                extracted_files.append({
                    "name": file_name.split("/")[-1],
                    "data": file_data
                })

    return extracted_files


folder_1a_files = extract_pdfs_from_zip(zip_1a)
folder_2a_files = extract_pdfs_from_zip(zip_2a)

all_uploaded_pdfs = folder_1a_files + folder_2a_files

# ==============================
# FIND ATTACHMENTS BASED ON PAN
# ==============================
def find_attachments(pan_no):
    matched_files = []

    for file in all_uploaded_pdfs:
        filename = file["name"]
        name_part = filename.split("_")[0].strip()

        if name_part == pan_no:
            matched_files.append(file)

    return matched_files


# ==============================
# SEND EMAIL FUNCTION
# ==============================
def send_emails(df):
    sent = 0
    skipped = 0
    failed = 0

    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(sender_email, sender_password)
    except Exception as e:
        st.error(f"SMTP Login Failed: {e}")
        return

    progress_bar = st.progress(0)
    log_area = st.empty()

    total_rows = len(df)

    for index, row in df.iterrows():
        pan = str(row['PAN']).strip()
        name = row['Name']
        email = row['Mail']

        files_found = find_attachments(pan)

        if not files_found:
            skipped += 1
            log_area.write(f"⚠ SKIPPED: {name} ({pan})")
            progress_bar.progress((index + 1) / total_rows)
            continue

        try:
            msg = MIMEMultipart()
            msg['From'] = sender_email
            msg['To'] = email
            msg['Subject'] = f"Official Document(s) for {name}"

            body = BODY_TEMPLATE.format(Name=name, PAN=pan)
            msg.attach(MIMEText(body, 'plain'))

            for file in files_found:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(file["data"])

                encoders.encode_base64(part)
                part.add_header(
                    "Content-Disposition",
                    f"attachment; filename={file['name']}"
                )
                msg.attach(part)

            server.send_message(msg)
            sent += 1
            log_area.write(f"✅ Sent to {name}")
            time.sleep(1.5)

        except Exception as e:
            failed += 1
            log_area.write(f"❌ FAILED for {name}: {e}")

        progress_bar.progress((index + 1) / total_rows)

    server.quit()

    st.success("Process Completed")
    st.write(f"Total Employees : {len(df)}")
    st.write(f"Sent            : {sent}")
    st.write(f"Skipped         : {skipped}")
    st.write(f"Failed          : {failed}")


# ==============================
# MAIN EXECUTION
# ==============================
if uploaded_excel is not None:

    df = pd.read_excel(uploaded_excel)

    required_columns = {"PAN", "Name", "Mail"}

    if not required_columns.issubset(set(df.columns)):
        st.error("Excel must contain columns: PAN, Name, Mail")
    else:
        st.success("Excel loaded successfully")
        st.dataframe(df.head())

        if st.button("🚀 Start Sending Emails"):
            if not sender_email or not sender_password:
                st.warning("Please enter Gmail credentials")
            elif not zip_1a and not zip_2a:
                st.warning("Please upload at least one ZIP folder (1A or 2A)")
            else:

                send_emails(df)


