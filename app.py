import base64, requests, zipfile, datetime as dt
import streamlit as st
import tax_report as tx

st.set_page_config(page_title="Tax Compliance Report", layout="centered")
st.title("Test Equity - OS - Tax Compliance Specialist Report")

if st.secrets.get("APP_PASSWORD") and st.text_input("Password", type="password") != st.secrets["APP_PASSWORD"]:
    st.stop()

GH = st.secrets.get("GITHUB", {})
if GH:
    _h = {"Authorization": f"Bearer {GH['token']}"}
    _u = f"https://api.github.com/repos/{GH['repo']}/contents/{GH['path']}"
    _p = {"ref": GH.get("branch", "main")}
    meta = requests.get(_u, headers=_h, params=_p, timeout=60)
    if meta.status_code != 200:
        st.error(f"GitHub replied {meta.status_code}: {meta.json().get('message')}. Check the token, repo name and file name in Secrets.")
        st.stop()
    sha = meta.json()["sha"]

    @st.cache_data(max_entries=1, show_spinner="Loading the report...")
    def fetch(_sha):
        return requests.get(_u, headers={**_h, "Accept": "application/vnd.github.raw+json"}, params=_p, timeout=300).content
    master = fetch(sha)
else:
    up = st.file_uploader("Current report (.xlsx)", type="xlsx"); master, sha = (up.getvalue() if up else None), None

files = st.file_uploader("The two agents' daily workbooks (select both together)", type="xlsx", accept_multiple_files=True)
if master and files:
    ctx = tx.learn(master); agents = ctx["agents"]
    dates = [tx.latest_date(f) for f in files]
    day = st.date_input("Date to load (each workbook holds many days; only this date is added)", value=max([d for d in dates if d] or [dt.date.today()]))
    jobs = []
    for f in files:
        g = tx.guess_agent(tx._load(zipfile.ZipFile(f))[0], agents)
        a = st.selectbox(f"Agent for {f.name}", agents, index=agents.index(g) if g in agents else 0, key=f.name,
                         format_func=lambda s: s.strip())
        jobs.append((f, a, day))
    if st.button("Update report", type="primary"):
        try:
            with st.spinner("Updating..."):
                new, log = tx.update_report(master, jobs)
        except ValueError as e:
            st.error(str(e)); st.stop()
        for line in log: (st.warning if "WARNING" in line else st.success if "+" in line else st.write)(line)
        if GH:
            r = requests.put(_u, headers=_h, timeout=600, json={"message": f"EOD update {day}", "content": base64.b64encode(new).decode(),
                                                                "sha": sha, "branch": GH.get("branch", "main")})
            st.info("Saved back to GitHub." if r.status_code in (200, 201) else f"Could not save to GitHub ({r.status_code}). Download the file below and tell the report owner.")
        st.download_button("Download updated report", new, f"Test_Equity_-_OS_-_Tax_Compliance_Specialist_{day:%d-%b-%Y}.xlsx")
