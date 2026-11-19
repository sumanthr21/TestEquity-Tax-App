# Tax Compliance Specialist report automation
Public code repo: app.py, tax_report.py, requirements.txt, README.md (NO report, NO token)
Private data repo: report.xlsx (latest report, about 3 MB)
Streamlit Secrets:
APP_PASSWORD = "choose-a-strong-password"
[GITHUB]
token = "github_pat_..."   # fine-grained, ONLY the private data repo, Contents: Read and write
repo = "yourname/tax-compliance-data"
path = "report.xlsx"
branch = "main"
Daily: MIS selects BOTH agents' workbooks together, confirms the date and each agent, clicks Update report.
