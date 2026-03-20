# HR Guru ATS v3 — File Upload Edition

No email/IMAP required. Works with Zoho free tier.

## How it works
1. You run the app on your computer (or a server)
2. Share the upload link with your team: http://[your-ip]:5000/upload
3. Recruiters fill the Excel template and upload it along with CVs
4. You view everything in the admin dashboard: http://localhost:5000

## Setup

### Install
```bash
cd ats3
pip install -r requirements.txt
```

### Run
```bash
python app.py
```

- Admin dashboard → http://localhost:5000
- Recruiter upload page → http://localhost:5000/upload

## Sharing with your team (same office network)
Find your local IP address:
- Windows: run `ipconfig` in Command Prompt → look for IPv4 Address (e.g. 192.168.1.10)
- Mac/Linux: run `ifconfig` or `ip addr`

Share http://192.168.1.10:5000/upload with your team.
They can open this on any device on the same WiFi.

## Excel template
Recruiters download the template from the upload page.
Column headers are flexible — the app maps common variations automatically.
They can also use their own format as long as column headers are recognisable.

## CV matching
Name CV files with the candidate's name: John_Smith.pdf, Priya_Sharma.docx
The app matches by looking for the candidate's name in the filename.

## Data
- Database: ats3/ats.db — back up regularly
- CVs: ats3/cvs/ — back up regularly
