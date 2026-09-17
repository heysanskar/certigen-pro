🎓 CertiGen Pro<br>
Professional Certificate Generation, Verification & Distribution Platform<br>
CertiGen Pro is a Streamlit-based certificate generation platform that automates the creation, personalization, verification, and distribution of digital certificates.<br>
The application allows users to upload Microsoft Word (.docx) or PowerPoint (.pptx) certificate templates containing dynamic placeholders such as {{name}}, {{designation}}, {{organization}}, {{topic}}, {{serial}}, {{date}}, and {{authcode}}.<br>

The system replaces these placeholders with recipient information, generates unique serial numbers and authentication codes, optionally embeds QR verification codes, converts certificates into PDF format, applies PDF permissions, stores audit records, and can automatically email certificates to recipients.<br>

<br>
🚀 Features<br>
📄 Upload Word (.docx) certificate templates<br>
📊 Upload PowerPoint (.pptx) certificate templates<br>
✏️ Dynamic placeholder replacement using {{}} syntax<br>
👤 Generate certificates for individual recipients<br>
📥 Bulk certificate generation using Excel<br>
🏷️ Support for recipient name, designation, organization, and topic<br>
🔢 Automatic serial number generation<br>
🔐 Unique authentication hash generation<br>
📱 QR code generation for certificate verification<br>
🔒 PDF permission protection<br>
📧 Automatic certificate email delivery<br>
📨 Custom email subject support<br>
💬 Custom email body support<br>
📋 CC recipient support<br>
📊 Certificate generation analytics<br>
📈 Email delivery statistics<br>
📥 Export audit reports as CSV<br>
🔍 Certificate verification using serial number and authentication code<br>
🗂️ Event-based certificate records<br>
🛡️ Login-protected application interface<br>
📝 Application logging<br>
💾 Local record storage<br>
<br>
🧠 How It Works<br>
The application follows this workflow.<br>
Certificate Template
        │
        ▼
Recipient Information
        │
        ▼
{{placeholder}} Replacement
        │
        ▼
Serial Number Generation
        │
        ▼
Authentication Hash
        │
        ▼
QR Code Generation
        │
        ▼
DOCX / PPTX Processing
        │
        ▼
PDF Conversion
        │
        ▼
PDF Permission Protection
        │
        ▼
Email Distribution
        │
        ▼
Audit Record Storage
        │
        ▼
Certificate Verification

<br>
📝 Template Placeholder System<br>
CertiGen Pro uses a simple {{placeholder}} system.<br>
You can create a certificate template in Microsoft Word or PowerPoint and place placeholders wherever personalized information should appear.<br>

Supported Placeholders<br>
{{name}}
{{designation}}
{{organization}}
{{topic}}
{{serial}}
{{date}}
{{authcode}}

For example, a certificate template can contain the following.<br>
This certificate is proudly presented to

{{name}}

for successfully participating as

{{designation}}

from

{{organization}}

in

{{topic}}

Certificate ID: {{serial}}

Issued on: {{date}}

Verification Code: {{authcode}}

When the certificate is generated, the placeholders are automatically replaced with the recipient's information.<br>
<br>
📱 QR Verification<br>
CertiGen Pro can optionally generate a QR code for every certificate.<br>
The QR code contains the certificate serial number and authentication code.<br>

Serial Number | Authentication Code

Example.<br>
TECH-170926-01|a81f92c3d4

The generated QR code can be placed into the certificate template using the following marker.<br>
[[QR]]

The system replaces the [[QR]] marker with the generated QR image.<br>
<br>
🔢 Serial Number Generation<br>
Every certificate receives a unique serial number based on the event code, generation date, and sequence number.<br>
Example.<br>

TECH-170926-01
TECH-170926-02
TECH-170926-03

The sequence automatically continues from the existing event records.<br>
The application also provides a reset option for starting the sequence again.<br>

<br>
🔐 Authentication System<br>
Each certificate receives an authentication code generated using SHA-256 hashing.<br>
The authentication code is derived from the serial number and recipient name.<br>

Serial Number + Recipient Name

Example.<br>
Serial:
TECH-170926-01

Authentication Code:
a81f92c3d4

These values are also stored in the event audit records.<br>
<br>
🔍 Certificate Verification<br>
The Audit Center allows users to verify generated certificates.<br>
The user provides the following information.<br>

System Serial String
Verification Hash Code

The application searches the stored event records and checks whether both values match.<br>
If a matching record is found, the certificate identity is displayed.<br>

<br>
📄 Word Certificate Generation<br>
CertiGen Pro supports Microsoft Word .docx templates.<br>
The application loads the Word template, searches for supported placeholders, replaces them with recipient information, attempts to preserve font styling, processes standard paragraphs and table content, inserts QR codes when enabled, and converts the generated document to PDF.<br>

<br>
📊 PowerPoint Certificate Generation<br>
CertiGen Pro also supports PowerPoint .pptx templates.<br>
The application processes slide text, text runs, tables, placeholder values, font information, font size, font colors, bold and italic formatting, and QR code placement.<br>

PowerPoint templates are converted to PDF using Microsoft PowerPoint on Windows.<br>

<br>
📥 Single Certificate Generation<br>
The application supports individual certificate generation.<br>
Enter the event name, event code, issue date, recipient name, recipient email, recipient designation, recipient organization, and topic.<br>

Then upload your .docx or .pptx certificate template.<br>

The system generates and processes the certificate automatically.<br>

<br>
📊 Bulk Certificate Generation<br>
CertiGen Pro supports bulk certificate generation using Excel files.<br>
Upload an .xlsx manifest containing columns such as the following.<br>

Name
Email
Designation
Organization
Topic

Example Excel structure.<br>
Name	Email	Designation	Organization	Topic
Rahul Sharma	rahul@example.com	Speaker	ABC Technologies	AI
Priya Singh	priya@example.com	Participant	XYZ University	Machine Learning

The application processes every valid row and generates an individual certificate.<br>
<br>
📧 Email Distribution<br>
Generated certificates can be automatically sent to recipients through SMTP.<br>
The application supports recipient email addresses, CC recipients, custom email subjects, custom email bodies, certificate attachments, serial number variables, and authentication code variables.<br>

Supported Email Variables<br>
{name}
{event}
{serial}
{auth}

Example subject.<br>
Congratulations {name} - {event}

Example email body.<br>
Dear {name},

Congratulations on participating in {event}.

Your certificate is attached to this email.

Certificate ID: {serial}
Verification Code: {auth}

Regards,
CertiGen Pro

<br>
🔒 PDF Security<br>
The application attempts to convert generated Word and PowerPoint certificates into PDF format.<br>
After successful PDF generation, the application applies PDF permission settings using pypdf.<br>

The objective is to restrict unauthorized modification of generated PDF certificates.<br>

Certificate generation is halted if secure PDF generation fails, preventing distribution of an editable template when the expected PDF output cannot be produced.<br>

<br>
📊 Analytics Dashboard<br>
The Analytics Reports section provides operational information about generated certificates.<br>
It tracks total certificates generated, successfully delivered emails, failed email deliveries, certificates that were not emailed, event records, and email delivery distribution.<br>

The application also provides a visual email-status chart using Plotly.<br>

<br>
📥 Audit Report Export<br>
Certificate records can be exported as CSV files.<br>
The audit records include event, name, email, designation, organization, topic, serial number, authentication code, date, and email status.<br>

This provides a persistent record of certificate generation and distribution.<br>

<br>
🗂️ Project Structure<br>
certigen-pro/
│
├── app.py
├── requirements.txt
├── README.md
├── .gitignore
├── LICENSE
│
├── system/
│   └── users.json
│
├── records/
│   └── event-code.csv
│
└── certificates/
    └── generated-certificates.pdf

The system, records, and certificates directories are generated and used locally by the application.<br>
Sensitive and generated files should not be committed to a public repository.<br>

<br>
🛠️ Tech Stack<br>
Technology	Purpose
Python	Application development
Streamlit	Web application interface
python-docx	Word document processing
python-pptx	PowerPoint processing
Pandas	Excel and CSV processing
qrcode	QR code generation
Pillow	Image processing
docx2pdf	Word to PDF conversion
comtypes	PowerPoint automation on Windows
pypdf	PDF processing and permissions
Plotly	Analytics visualization
SMTP	Email delivery
hashlib	Authentication hash generation
JSON	Local user authentication
Logging	Application logging

<br>
⚙️ Installation<br>
📋 Prerequisites<br>
Make sure you have the following installed.<br>
Python 3.9 or newer<br>
Git<br>
Microsoft Word for DOCX-to-PDF conversion<br>
Microsoft PowerPoint for PPTX-to-PDF conversion on Windows<br>
A Gmail account or another SMTP-enabled email account for certificate delivery<br>
<br>
1. Clone the Repository<br>
git clone https://github.com/YOUR_USERNAME/certigen-pro.git
cd certigen-pro

<br>
2. Create a Virtual Environment<br>
Windows<br>
python -m venv venv
venv\Scripts\activate

macOS / Linux<br>
python3 -m venv venv
source venv/bin/activate

<br>
3. Install Dependencies<br>
pip install -r requirements.txt

<br>
4. Run the Application<br>
streamlit run app.py

The application will open in your browser.<br>
<br>
🔐 Login<br>
On the first run, CertiGen Pro creates the following local file.<br>
system/users.json

The application currently creates a default local account.<br>
Username:
admin

Password:
password123

Change the default password before using the application in a real environment.<br>
Do not commit system/users.json to GitHub.<br>

<br>
📧 Gmail SMTP Configuration<br>
If you want to send certificates through Gmail, use a Gmail App Password rather than your normal Gmail account password.<br>
The application requires an SMTP user email and SMTP App Password.<br>

The SMTP connection uses the following configuration.<br>

smtp.gmail.com
Port: 465
SSL

Never commit your email password or App Password to GitHub.<br>
<br>
🛡️ Security & Privacy<br>
CertiGen Pro processes certificate data locally.<br>
Generated records may contain recipient names, email addresses, organization names, certificate IDs, and authentication codes.<br>

Therefore, do not upload real recipient data to a public repository.<br>

Do not commit generated certificates, users.json, SMTP passwords, or App Passwords to GitHub.<br>

Keep generated records and certificates in local storage.<br>

Use secure environment-based credentials in production environments.<br>

<br>
⚠️ Important Conversion Requirement<br>
DOCX-to-PDF and PPTX-to-PDF conversion depends on the local environment.<br>
PowerPoint PDF conversion requires Windows and Microsoft PowerPoint.<br>

The application does not provide native PPTX-to-PDF conversion on systems without the required Microsoft Office environment.<br>

<br>
🎯 Project Objective<br>
The objective of CertiGen Pro is to automate the complete certificate lifecycle.<br>
Instead of manually editing certificates one by one, the system combines template processing, recipient data, dynamic placeholder replacement, serial numbering, authentication hashing, QR generation, PDF generation, PDF protection, email distribution, audit logging, and certificate verification.<br>

The complete workflow can be represented as follows.<br>

Template
   ↓
Recipient Data
   ↓
Dynamic Placeholder Replacement
   ↓
Serial Number
   ↓
Authentication Hash
   ↓
QR Code
   ↓
PDF Generation
   ↓
PDF Protection
   ↓
Email Distribution
   ↓
Audit Logging
   ↓
Certificate Verification

This makes the certificate generation process faster, more consistent, and easier to audit.<br>
<br>
⭐ Key Highlights<br>
🎓 Automated certificate generation<br>
📄 Word and PowerPoint template support<br>

🔄 Dynamic {{}} placeholder replacement<br>

📊 Excel bulk processing<br>

🔢 Automatic serial numbering<br>

🔐 Authentication hash generation<br>

📱 QR-enabled certificates<br>

🔒 PDF permission protection<br>

📧 Automated email distribution<br>

💬 Custom email templates<br>

📊 Analytics dashboard<br>

📋 Audit logging<br>

🔍 Certificate verification<br>

📥 CSV report export<br>

🛡️ Login-protected application<br>

<br>
🔮 Future Improvements<br>
🔎 Advanced certificate search and filtering<br>
🗑️ Certificate deletion and management<br>

🔄 Advanced duplicate detection<br>

🌐 Online certificate verification portal<br>

📱 Mobile-friendly certificate verification<br>

🔐 Stronger user authentication<br>

👥 Multiple user roles and permissions<br>

🔑 Secure environment-based credential management<br>

🗃️ Database-backed record management<br>

☁️ Optional cloud storage<br>

📊 Advanced analytics and reporting<br>

📜 Certificate revocation system<br>

🔗 Public verification URLs<br>

📱 QR codes linked to online verification pages<br>

🖨️ Batch printing support<br>

🎨 Template preview functionality<br>

<br>
📚 Key Concepts Demonstrated<br>
This project demonstrates practical implementation of document automation, Microsoft Word processing, Microsoft PowerPoint processing, dynamic template placeholders, bulk Excel processing, PDF generation, PDF permission management, QR code generation, SHA-256 hashing, SMTP email automation, Streamlit application development, session state management, local authentication, CSV audit logging, data visualization, and automated document distribution.<br>
<br>
👨‍💻 Author<br>
Sanskar Aman<br>
GitHub: https://github.com/heysanskar<br>

LinkedIn: https://www.linkedin.com/in/sanskar-a-881719248/<br>

<br>
⭐ Support<br>
If you find this project useful, consider giving the repository a star on GitHub.<br>
<br>
📄 License<br>
This project is licensed under the MIT License.<br>
See the LICENSE file for details.<br>
