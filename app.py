# ═══════════════════════════════════════════════════════════════════════════════
#  Berea College — Administrative Support Request System  v3.0
#  Office of the Provost & Dean of Faculty Development
#
#  FEATURES: Multi-user auth · 13-branch form · Admin dashboard ·
#  Supervisor portal · Email notifications · Notification preferences ·
#  Google Sheets persistent storage (session-state fallback)
#
#  BRAND: 2026 Official Standards
#  Berea Blue #004175 | Chartreuse #D2DE26 | Light Blue #49B4EF
#  Fonts: Barlow (body) · Barlow Condensed (headers) · Newsreader (headings)
# ═══════════════════════════════════════════════════════════════════════════════

import streamlit as st
import pandas as pd
import json, smtplib
from datetime import datetime, date
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

st.set_page_config(
    page_title="Support Requests · Berea College Provost's Office",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded",
    menu_items={"Get help":None,"Report a bug":None,
                "About":"Berea College Provost's Office — Support Request System v3.0"}
)

BLUE="#004175"; CHART="#D2DE26"; LBLUE="#49B4EF"; BLACK="#000000"; WHITE="#FFFFFF"

st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Barlow:wght@400;500;600;700&family=Barlow+Condensed:wght@600;700&family=Newsreader:wght@400;600&display=swap');
html,[class*="css"]{{font-family:'Barlow',sans-serif;}}
#MainMenu,footer,[data-testid="stToolbar"]{{visibility:hidden;height:0;}}
header{{visibility:hidden;}}
.block-container{{padding-top:1rem;padding-bottom:3rem;}}
.banner{{background:{BLUE};border-bottom:5px solid {CHART};border-radius:10px;padding:1.1rem 1.5rem .9rem;margin-bottom:1.1rem;}}
.banner h1{{font-family:'Newsreader',serif;color:{WHITE};font-size:1.5rem;font-weight:600;margin:0 0 2px;}}
.banner p{{font-family:'Barlow Condensed',sans-serif;color:{LBLUE};font-size:.72rem;font-weight:700;letter-spacing:.1em;text-transform:uppercase;margin:0;}}
.sec{{background:{BLUE};color:{WHITE};font-family:'Barlow Condensed',sans-serif;font-weight:700;font-size:.82rem;letter-spacing:.1em;text-transform:uppercase;padding:6px 12px;border-radius:5px;margin:1.4rem 0 .7rem;border-left:5px solid {CHART};}}
.hr{{border:none;border-top:1px solid #cdd6e0;margin:1.1rem 0;}}
.note-card{{background:{WHITE};border:0.5px solid #cdd6e0;border-left:4px solid {BLUE};border-radius:0 8px 8px 0;padding:9px 13px;margin-bottom:7px;font-size:.84rem;}}
.note-card.tiff-note{{border-left-color:{CHART};background:#fafdf0;}}
.note-meta{{font-size:.71rem;color:#5a7a96;margin-bottom:3px;font-weight:600;}}
.note-meta.tiff{{color:#3B6D11;}}
.sb{{padding:3px 9px;border-radius:10px;font-size:.71rem;font-weight:700;font-family:'Barlow Condensed',sans-serif;letter-spacing:.04em;display:inline-block;}}
.s-new{{background:#e8f1fb;color:{BLUE};}} .s-review{{background:#eef2ff;color:#1a3a6e;}}
.s-prog{{background:#fff3cd;color:#7a4f00;}} .s-await{{background:#fde8e8;color:#9b1c1c;}}
.s-done{{background:#eaf3de;color:#27500A;}} .s-hold{{background:#f1efe8;color:#444;}}
.urg-u{{background:#fde8e8;color:#9b1c1c;}} .urg-h{{background:#fff3cd;color:#7a4f00;}}
.urg-s{{background:#e6f0fa;color:{BLUE};}} .urg-l{{background:#f0fdf4;color:#166534;}}
.ro-notice{{background:#e8f1fb;border-left:4px solid {LBLUE};border-radius:0 6px 6px 0;padding:7px 11px;font-size:.79rem;color:#003060;margin:.7rem 0;}}
.mc{{background:{WHITE};border:0.5px solid #cdd6e0;border-radius:10px;padding:11px 13px;text-align:center;}}
.mn{{font-size:1.7rem;font-weight:700;line-height:1;}} .mn.bl{{color:{BLUE};}} .mn.am{{color:#854F0B;}} .mn.gr{{color:#3B6D11;}}
.ml{{font-size:.71rem;color:#5a7a96;font-weight:700;font-family:'Barlow Condensed',sans-serif;letter-spacing:.04em;text-transform:uppercase;margin-top:3px;}}
.ref-code{{font-family:'Barlow Condensed',monospace;font-size:.84rem;font-weight:700;color:{BLUE};letter-spacing:.05em;}}
.conf-flag{{background:#fde8e8;color:#9b1c1c;font-size:.7rem;font-weight:700;padding:2px 6px;border-radius:8px;font-family:'Barlow Condensed',sans-serif;}}
div[data-testid="stButton"]>button{{background:{BLUE}!important;color:{WHITE}!important;font-family:'Barlow Condensed',sans-serif!important;font-weight:700!important;font-size:.88rem!important;letter-spacing:.06em!important;text-transform:uppercase!important;border:none!important;border-bottom:3px solid {CHART}!important;border-radius:6px!important;}}
div[data-testid="stButton"]>button:hover{{background:#003060!important;}}
[data-testid="stSidebar"]{{background:#f5f8fc;border-right:1px solid #cdd6e0;}}
</style>
""", unsafe_allow_html=True)

# ── Config ──────────────────────────────────────────────────────────────────
def _s(k,fb=""): 
    try: return st.secrets[k]
    except: return fb

ADMIN_PASSWORD = _s("ADMIN_PASSWORD","admin2025!")
ADMIN_EMAIL    = _s("ADMIN_EMAIL","")
ADMIN_NAME     = _s("ADMIN_NAME","Tiff (Provost's Office)")
GMAIL_USER     = _s("GMAIL_USER","")
GMAIL_PASS     = _s("GMAIL_APP_PASSWORD","")
SHEET_ID       = _s("GOOGLE_SHEET_ID","").strip()
SA_JSON        = _s("GOOGLE_SERVICE_ACCOUNT","").strip()

def load_supervisors():
    try:
        raw = st.secrets.get("supervisors",{})
        return {info.get("display_name",k): dict(info) for k,info in raw.items()}
    except: return {}

SUPERVISORS = load_supervisors()

# ── Data Layer ───────────────────────────────────────────────────────────────
class DataLayer:
    def __init__(self):
        self.ok=False; self._wb=None; self.err=""; self.err_stage=""; self._try()
    def _try(self):
        if not SHEET_ID or not SA_JSON:
            self.err="GOOGLE_SHEET_ID or GOOGLE_SERVICE_ACCOUNT is missing from secrets."
            self.err_stage="secrets"
            return
        try:
            import gspread
            from google.oauth2.service_account import Credentials
        except Exception as e:
            self.err=f"Required library not installed: {e}"; self.err_stage="import"; return
        try:
            info=json.loads(SA_JSON)
        except Exception as e:
            self.err=f"The service-account JSON couldn't be read: {e}"; self.err_stage="json"; return
        try:
            sc=["https://spreadsheets.google.com/feeds","https://www.googleapis.com/auth/drive"]
            c=Credentials.from_service_account_info(info,scopes=sc)
            gc=gspread.authorize(c)
        except Exception as e:
            self.err=f"Could not authenticate with the service account: {e}"; self.err_stage="auth"; return
        try:
            self._wb=gc.open_by_key(SHEET_ID); self.ok=True
        except Exception as e:
            self.err=f"Authenticated OK, but couldn't open the Sheet: {e}"; self.err_stage="open"; return
    def _ws(self,name,hdrs):
        try: return self._wb.worksheet(name)
        except:
            ws=self._wb.add_worksheet(title=name,rows=2000,cols=max(len(hdrs),10))
            if hdrs: ws.append_row(hdrs)
            return ws
    SH=["ref_number","submitted_at","Your Name","Your Role","Contact Info",
        "Deadline","Urgency","Confidential","Request Type","Status","Details JSON",
        "Notification Email","Notify on Status Change","Notify on New Note"]
    def save_submission(self,data,ref):
        d=data.copy(); d["ref_number"]=ref; d.setdefault("Status","New")
        if self.ok:
            ws=self._ws("submissions",self.SH)
            det={k:v for k,v in d.items() if k not in self.SH}
            row=[str(d.get(h,"")) for h in self.SH[:-1]]+[json.dumps(det)]
            ws.append_row(row)
        else:
            if "submissions" not in st.session_state: st.session_state.submissions=[]
            st.session_state.submissions.append(d)
    def get_submissions(self,name=None):
        if self.ok:
            try:
                ws=self._ws("submissions",self.SH); rows=ws.get_all_records()
                return [r for r in rows if not name or r.get("Your Name","")==name]
            except: return []
        subs=list(st.session_state.get("submissions",[]))
        return [s for s in subs if not name or s.get("Your Name","")==name]
    def update_status(self,ref,status):
        if self.ok:
            try:
                ws=self._ws("submissions",self.SH); cell=ws.find(ref)
                if cell:
                    h=ws.row_values(1); ws.update_cell(cell.row,h.index("Status")+1,status)
            except: pass
        else:
            for s in st.session_state.get("submissions",[]):
                if s.get("ref_number")==ref: s["Status"]=status
    NH=["ref_number","author","author_role","timestamp","text"]
    def add_note(self,ref,author,role,text):
        n={"ref_number":ref,"author":author,"author_role":role,
           "timestamp":datetime.now().strftime("%b %d %Y %I:%M %p"),"text":text}
        if self.ok:
            ws=self._ws("notes",self.NH); ws.append_row([n[h] for h in self.NH])
        else:
            k=f"notes_{ref}"
            if k not in st.session_state: st.session_state[k]=[]
            st.session_state[k].append(n)
        return n
    def get_notes(self,ref):
        if self.ok:
            try:
                ws=self._ws("notes",self.NH)
                return [r for r in ws.get_all_records() if r.get("ref_number")==ref]
            except: return []
        return list(st.session_state.get(f"notes_{ref}",[]))
    PH=["username","on_new_note","on_status_change","on_new_submission","on_any_note"]
    PD={"on_new_note":True,"on_status_change":True,"on_new_submission":True,"on_any_note":True}
    def get_prefs(self,user):
        if self.ok:
            try:
                ws=self._ws("notification_prefs",self.PH)
                for r in ws.get_all_records():
                    if r.get("username")==user:
                        return {k:str(v).lower() not in("false","0","no","")
                                for k,v in r.items() if k!="username"}
            except: pass
        return st.session_state.get(f"prefs_{user}",dict(self.PD))
    def save_prefs(self,user,prefs):
        prefs["username"]=user
        if self.ok:
            try:
                ws=self._ws("notification_prefs",self.PH)
                rows=ws.get_all_records(); row=[str(prefs.get(h,"")) for h in self.PH]
                for i,r in enumerate(rows):
                    if r.get("username")==user:
                        for j,v in enumerate(row): ws.update_cell(i+2,j+1,v)
                        return
                ws.append_row(row)
            except: pass
        st.session_state[f"prefs_{user}"]=prefs
    # ── PERSONAL TASKS (Tiff's own to-dos) ─────────────────────────────────────
    TH=["task_id","created_at","Title","Category","Priority","Deadline",
        "Status","Recurrence","Details","source_ref","Visibility"]
    def save_task(self,task):
        if self.ok:
            ws=self._ws("tasks",self.TH); ws.append_row([str(task.get(h,"")) for h in self.TH])
        else:
            if "tasks" not in st.session_state: st.session_state.tasks=[]
            st.session_state.tasks.append(task)
    def get_tasks(self):
        if self.ok:
            try: return self._ws("tasks",self.TH).get_all_records()
            except: return []
        return list(st.session_state.get("tasks",[]))
    def update_task(self,task_id,field,value):
        if self.ok:
            try:
                ws=self._ws("tasks",self.TH); cell=ws.find(task_id)
                if cell:
                    h=ws.row_values(1); ws.update_cell(cell.row,h.index(field)+1,str(value))
            except: pass
        else:
            for t in st.session_state.get("tasks",[]):
                if t.get("task_id")==task_id: t[field]=value
    def delete_task(self,task_id):
        if self.ok:
            try:
                ws=self._ws("tasks",self.TH); cell=ws.find(task_id)
                if cell: ws.delete_rows(cell.row)
            except: pass
        else:
            st.session_state.tasks=[t for t in st.session_state.get("tasks",[]) if t.get("task_id")!=task_id]

@st.cache_resource
def get_db(): return DataLayer()
DB=get_db()

# ── Notifications ────────────────────────────────────────────────────────────
def _email(to,subj,body):
    if not all([GMAIL_USER,GMAIL_PASS,to]): return
    try:
        msg=MIMEMultipart(); msg["From"]=GMAIL_USER; msg["To"]=to; msg["Subject"]=subj
        msg.attach(MIMEText(body,"plain"))
        with smtplib.SMTP_SSL("smtp.gmail.com",465,timeout=10) as s:
            s.login(GMAIL_USER,GMAIL_PASS); s.sendmail(GMAIL_USER,to,msg.as_string())
    except: pass

def notify_new_submission(data,ref):
    p=DB.get_prefs("__admin__")
    if not p.get("on_new_submission",True) or not ADMIN_EMAIL: return
    conf="Yes" in str(data.get("Confidential",""))
    subj=f"{'[CONFIDENTIAL] ' if conf else ''}New request {ref} — {data.get('Request Type','').split('  ')[-1]}"
    body=(f"New support request submitted.\n\nFrom: {data.get('Your Name','')} ({data.get('Your Role','')})\n"
          f"Ref: {ref}\nType: {data.get('Request Type','')}\nUrgency: {data.get('Urgency','')}\n"
          f"Deadline: {data.get('Deadline','')}\nSubmitted: {data.get('submitted_at','')}\n")
    _email(ADMIN_EMAIL,subj,body)

def notify_status_change(ref,new_status,sub_name,sub_email):
    if not sub_email: return
    p=DB.get_prefs(sub_name)
    if not p.get("on_status_change",True): return
    _email(sub_email,f"Your request {ref} — status updated to: {new_status}",
           f"Your support request {ref} status has been updated to: {new_status}\n\nLog in to view details and add notes.\n")

def notify_accepted(ref,rtype,sub_name,sub_email):
    if not sub_email: return
    p=DB.get_prefs(sub_name)
    if not p.get("on_status_change",True): return
    _email(sub_email,f"Your request {ref} has been accepted",
           f"Good news — Tiff has accepted your support request and started working on it.\n\n"
           f"Reference: {ref}\nRequest: {rtype}\n\n"
           f"You can log in anytime to check its status or add a note if anything changes.\n")

def notify_declined(ref,rtype,sub_name,sub_email,reason=""):
    if not sub_email: return
    p=DB.get_prefs(sub_name)
    if not p.get("on_status_change",True): return
    body=(f"Thank you for your request. Unfortunately Tiff isn't able to take this one on right now.\n\n"
          f"Reference: {ref}\nRequest: {rtype}\n")
    if reason.strip(): body+=f"\nNote: {reason.strip()}\n"
    body+="\nPlease feel free to reach out directly if you'd like to discuss alternatives.\n"
    _email(sub_email,f"Update on your request {ref}",body)

def notify_new_note(ref,author,text,recipient_name,recipient_email):
    if not recipient_email: return
    is_admin=(recipient_name=="__admin__")
    p=DB.get_prefs(recipient_name)
    if not p.get("on_any_note" if is_admin else "on_new_note",True): return
    _email(recipient_email,f"New note on request {ref}",
           f"A new note has been added to request {ref}.\n\nFrom: {author}\n\nNote:\n{text}\n\nLog in to reply.\n")

def notify_task_comment(task_id,task_title,author,text):
    # A supervisor commented on one of Tiff's tasks -> notify Tiff (admin).
    if not ADMIN_EMAIL: return
    p=DB.get_prefs("__admin__")
    if not p.get("on_any_note",True): return
    _email(ADMIN_EMAIL,f"New comment on your task: {task_title}",
           f"A supervisor commented on one of your tasks.\n\nTask: {task_title}\nFrom: {author}\n\nComment:\n{text}\n\nLog in to view.\n")

# ── Session state ────────────────────────────────────────────────────────────
for k,v in {"user":None,"user_role":None,"submitted":False,"submission_data":{},
            "ref_number":"","view":"form"}.items():
    if k not in st.session_state: st.session_state[k]=v

def is_admin(): return st.session_state.user_role=="admin"
def is_sup(): return st.session_state.user_role=="supervisor"
def cur_user(): return st.session_state.user

# ── Helpers ──────────────────────────────────────────────────────────────────
REQUEST_TYPES=["— Select a request type —",
    "📅  Event Planning & Coordination",
    "🗓️  Meeting Scheduling & Calendar Management",
    "✈️  Travel & Logistics Coordination",
    "📝  Document Preparation or Editing",
    "📣  Communication Drafting (Emails, Letters, Memos, Announcements)",
    "🏛️  Room / Space Reservation Only",
    "🍽️  Catering & Hospitality",
    "👤  Guest or Visitor Coordination",
    "💰  Budget, Purchasing & Expense Processing",
    "🔍  Research & Data Compilation",
    "🎓  Faculty Development Programming Support",
    "📊  Presentation & Materials Preparation",
    "🗂️  Other / General Administrative Support"]
STATUS_OPTIONS=["New","Under Review","Accepted","In Progress","Awaiting Info","Complete","On Hold","Declined"]
S_CSS={"New":"s-new","Under Review":"s-review","Accepted":"s-prog","In Progress":"s-prog",
       "Awaiting Info":"s-await","Complete":"s-done","On Hold":"s-hold","Declined":"s-await"}

def banner(t,s="Berea College · Office of the Provost & Dean of Faculty Development"):
    st.markdown(f'<div class="banner"><h1>{t}</h1><p>{s}</p></div>',unsafe_allow_html=True)
def sec(l): st.markdown(f'<div class="sec">{l}</div>',unsafe_allow_html=True)
def hr(): st.markdown('<hr class="hr">',unsafe_allow_html=True)
def lbl(t): return f"{t} *"
def gen_ref(): return f"REQ-{datetime.now().strftime('%Y%m%d-%H%M%S')}"
def gen_task_id(): return f"TASK-{datetime.now().strftime('%Y%m%d-%H%M%S')}"
TASK_CATEGORIES=["Scheduling / calendar","Minutes / documentation","Committee support",
    "Event / logistics","Communications","Data / reporting","Systems / digital projects",
    "Follow-up / admin","Personal / professional dev","Other"]
TASK_PRIORITIES=["🔴 Urgent","🟠 High","🟡 Medium","🟢 Low"]
RECURRENCE_OPTIONS=["None — one-time","Daily","Weekly","Biweekly","Monthly","Quarterly","Annually"]
def next_occurrence(deadline_str, recurrence):
    """Compute the next deadline for a recurring task."""
    try:
        from datetime import timedelta
        d=datetime.strptime(deadline_str,"%Y-%m-%d").date()
    except: return deadline_str
    from datetime import timedelta
    deltas={"Daily":timedelta(days=1),"Weekly":timedelta(weeks=1),
            "Biweekly":timedelta(weeks=2),"Monthly":timedelta(days=30),
            "Quarterly":timedelta(days=91),"Annually":timedelta(days=365)}
    if recurrence in deltas: return (d+deltas[recurrence]).strftime("%Y-%m-%d")
    return deadline_str
def ucss(u): return("urg-u" if "Urgent" in u else "urg-h" if "High" in u else "urg-s" if "Standard" in u else "urg-l")
def friendly_status(s):
    # Supervisor-facing wording for each status.
    return {
        "New":"Submitted & pending confirmation of availability",
        "":"Submitted & pending confirmation of availability",
        "Under Review":"Under review by Tiff",
        "Accepted":"Confirmed & underway",
        "In Progress":"In progress",
        "Awaiting Info":"Awaiting your input",
        "Complete":"Complete",
        "On Hold":"On hold",
        "Declined":"Not able to take this on",
    }.get(s,s)
def summary(data,ref):
    lines=["BEREA COLLEGE — SUPPORT REQUEST","="*50,f"Ref: {ref}",
           f"Submitted: {data.get('submitted_at','')}",f"Status: {data.get('Status','New')}","="*50]
    for k,v in data.items():
        if k not in {"submitted_at","ref_number"}: lines.append(f"{k}: {v if v else '—'}")
    return "\n".join(lines)

# ── Sidebar ──────────────────────────────────────────────────────────────────
def sidebar():
    with st.sidebar:
        st.markdown(f'<div style="background:{BLUE};border-bottom:3px solid {CHART};border-radius:8px;'
                    f'padding:9px 13px;margin-bottom:12px;">'
                    f'<div style="font-family:\'Barlow Condensed\',sans-serif;font-weight:700;'
                    f'color:{WHITE};font-size:.88rem;letter-spacing:.06em;text-transform:uppercase;">🎓 Provost\'s Office</div>'
                    f'<div style="color:{LBLUE};font-size:.72rem;margin-top:2px;">{cur_user() or ""}</div></div>',
                    unsafe_allow_html=True)
        if is_admin():
            if st.button("📊  Admin dashboard",key="nb1",use_container_width=True):
                st.session_state.view="admin"; st.session_state.submitted=False; st.rerun()
            if st.button("✅  My tasks",key="nb1b",use_container_width=True):
                st.session_state.view="tasks"; st.session_state.submitted=False; st.rerun()
            if st.button("📝  Submit new request",key="nb2",use_container_width=True):
                st.session_state.view="form"; st.session_state.submitted=False; st.rerun()
        elif is_sup():
            if st.button("📋  My requests",key="nb3",use_container_width=True):
                st.session_state.view="portal"; st.session_state.submitted=False; st.rerun()
            if st.button("➕  Submit new request",key="nb4",use_container_width=True):
                st.session_state.view="form"; st.session_state.submitted=False; st.rerun()
        if st.session_state.user:
            st.markdown("---")
            if st.button("🔔  Notification settings",key="nb5",use_container_width=True):
                st.session_state.view="prefs"; st.rerun()
            if st.button("🔒  Log out",key="nb6",use_container_width=True):
                st.session_state.user=None; st.session_state.user_role=None
                st.session_state.view="form"; st.session_state.submitted=False; st.rerun()
        st.markdown(
            f'<div style="margin-top:12px;background:{"#eaf3de" if DB.ok else "#fff3cd"};'
            f'border-radius:6px;padding:7px 9px;font-size:.73rem;'
            f'color:{"#27500A" if DB.ok else "#7a4f00"};">'
            f'{"✅ Google Sheets storage active" if DB.ok else "⚠️ Session mode — data resets on restart. Configure Google Sheets for persistence."}'
            f'</div>',unsafe_allow_html=True)
        # ── Admin-only storage diagnostic ──────────────────────────────────
        if is_admin() and not DB.ok:
            hints={
                "secrets":"One or both secrets are missing. Add GOOGLE_SHEET_ID and GOOGLE_SERVICE_ACCOUNT in Settings → Secrets.",
                "import":"A required library isn't installed. Make sure requirements.txt includes gspread and google-auth.",
                "json":"The service-account JSON is incomplete or malformed. Re-copy the ENTIRE JSON file (from the opening { to the closing }) between triple quotes.",
                "auth":"The credentials were rejected. The JSON key may be truncated or from the wrong project. Re-download a fresh key and paste it whole.",
                "open":"Authentication worked, but the Sheet couldn't be opened. This is almost always ONE of: (1) the Sheet isn't shared with the service-account email (client_email in the JSON) as Editor, (2) the GOOGLE_SHEET_ID is wrong (should be only the code between /d/ and /edit), or (3) the Google Drive API isn't enabled for the project.",
            }
            with st.expander("🔎 Storage diagnostic (admin only)"):
                st.caption("Why the app is in session mode right now:")
                _err=getattr(DB,"err","") or "No error captured."
                _stage=getattr(DB,"err_stage","")
                st.code(_err,language=None)
                if _stage in hints:
                    st.markdown(f"**Most likely fix:** {hints[_stage]}")
                # Show which secrets the app can actually see (names/lengths only — never values).
                st.caption("What the app can see in your secrets (values hidden for safety):")
                def _peek(k):
                    try:
                        v=st.secrets.get(k,None)
                    except Exception:
                        v=None
                    if v is None: return f"❌ {k}: NOT FOUND (check the exact spelling)"
                    s=str(v)
                    return f"✅ {k}: present ({len(s)} characters)" if s.strip() else f"⚠️ {k}: found but EMPTY"
                st.code("\n".join([_peek("GOOGLE_SHEET_ID"),_peek("GOOGLE_SERVICE_ACCOUNT"),
                                   _peek("ADMIN_PASSWORD"),_peek("ADMIN_EMAIL")]),language=None)
                try:
                    sups=st.secrets.get("supervisors",{})
                    st.caption(f"Supervisor accounts detected: {len(sups)}")
                except Exception:
                    st.caption("Supervisor accounts detected: could not read [supervisors] section")
                st.caption("This panel is visible only to you (admin), and only while storage is not connected.")
                st.markdown("---")
                st.caption("Run a live connection test right now (bypasses any cached code):")
                if st.button("🔌 Test connection now",key="diag_test"):
                    stage="start"
                    try:
                        stage="reading secrets"
                        sid=str(st.secrets.get("GOOGLE_SHEET_ID","")).strip()
                        sa=str(st.secrets.get("GOOGLE_SERVICE_ACCOUNT",""))
                        if not sid or not sa:
                            st.error("Secrets missing at test time."); st.stop()
                        stage="importing libraries"
                        import gspread
                        from google.oauth2.service_account import Credentials
                        stage="parsing the JSON key"
                        info=json.loads(sa)
                        st.info(f"Service account email in your key:\n\n{info.get('client_email','(none found)')}")
                        stage="authenticating"
                        scopes=["https://spreadsheets.google.com/feeds","https://www.googleapis.com/auth/drive"]
                        creds=Credentials.from_service_account_info(info,scopes=scopes)
                        gc=gspread.authorize(creds)
                        stage="opening the Sheet by ID"
                        wb=gc.open_by_key(sid)
                        stage="reading the first worksheet"
                        title=wb.title
                        st.success(f"✅ SUCCESS — connected and opened the Sheet titled: “{title}”.\n\nReorder/reboot done. Storage should now work — reboot the app once more to switch out of session mode.")
                    except Exception as e:
                        st.error(f"❌ Failed while: **{stage}**\n\nError:\n\n{type(e).__name__}: {e}")
                        st.caption("Tell Claude the stage and error above and you'll get the exact fix.")

# ── Form branches ─────────────────────────────────────────────────────────────
def br_event():
    d={}; sec("📅 Event Planning & Coordination")
    d["Event Name"]=st.text_input(lbl("Event name or working title"),key="e_nm")
    d["Event Purpose"]=st.text_area(lbl("Purpose and goal"),placeholder="Describe the purpose, audience, and what success looks like.",key="e_pur")
    d["Event Type"]=st.selectbox(lbl("Type of event"),["— Select —","Lecture or speaker presentation","Workshop or training","Reception or social gathering","Award ceremony or recognition event","Retreat","Symposium or conference","Commencement or formal ceremony","Community or outreach event","Committee or working group meeting","Other"],key="e_typ")
    d["Proposed Date(s)"]=st.text_input(lbl("Proposed date or date range"),placeholder="Preferred date AND backup dates",key="e_dt")
    d["Proposed Time"]=st.text_input(lbl("Start and end time"),placeholder="e.g., 3:00 PM – 5:00 PM ET",key="e_tm")
    d["Location"]=st.selectbox(lbl("Location"),["— Select —","On campus — already identified","On campus — TBD, please assist","Off campus — venue identified","Off campus — venue TBD, please assist","Virtual (Teams, Zoom, or other)","Hybrid"],key="e_loc")
    d["Location Detail"]=st.text_input("If already identified, specify:",key="e_lcd")
    d["Expected Attendees"]=st.number_input(lbl("Expected attendees"),min_value=1,max_value=5000,value=20,key="e_ct")
    d["Audience"]=", ".join(st.multiselect(lbl("Intended audience"),["Berea College faculty","Berea College staff","Students","External guests","Community members","Board members or trustees","Donors or alumni","Other"],key="e_aud"))
    d["Catering"]=st.radio(lbl("Food or beverages?"),["Yes — specific needs","Yes — please advise","No","Not sure yet"],key="e_cat",horizontal=True)
    if "Yes" in str(d.get("Catering","")): d["Catering Details"]=st.text_area("Describe catering needs:",key="e_catd")
    d["Guest Speaker"]=st.radio(lbl("Guest speaker or presenter?"),["Yes — already confirmed","Yes — needs to be identified","No"],key="e_spk",horizontal=True)
    if "confirmed" in str(d.get("Guest Speaker","")): d["Speaker Contact"]=st.text_input("Speaker name and contact:",key="e_spkc")
    d["Event Needs"]=", ".join(st.multiselect(lbl("This event will require"),["A/V setup","Recording or livestreaming","Photography","Printed programs","Name tags / registration","Guest honoraria","Transportation or parking","Accessibility accommodations","Budget tracking"],key="e_nds"))
    d["Budget Status"]=st.selectbox(lbl("Budget status"),["— Select —","Yes — approved, account available","Yes — approved, account TBD","Pending approval","No budget — low/no-cost","Need help determining budget"],key="e_bud")
    d["Budget Detail"]=st.text_input("Budget amount and account/fund number:",key="e_budd")
    d["Additional Notes"]=st.text_area("Additional notes:",key="e_nt"); return d

def br_meeting():
    d={}; sec("🗓️ Meeting Scheduling & Calendar Management")
    d["Meeting Purpose"]=st.text_input(lbl("Purpose or topic"),placeholder="e.g., 'Fall Curriculum Review'",key="m_pur")
    d["Required Attendees"]=st.text_area(lbl("Who needs to attend?"),placeholder="List by name and/or role.",key="m_att")
    d["Optional Attendees"]=st.text_input("Optional attendees:",key="m_opt")
    d["Format"]=st.radio(lbl("Preferred format"),["In person","Virtual (Teams)","Virtual (Zoom)","Hybrid","No preference"],key="m_fmt",horizontal=True)
    d["Duration"]=st.selectbox(lbl("Length"),["— Select —","30 minutes","45 minutes","1 hour","1.5 hours","2 hours","More than 2 hours"],key="m_dur")
    d["Preferred Timeframe"]=st.text_area(lbl("Preferred timeframe"),placeholder="Target date, range, or recurring schedule. Note blackout dates.",key="m_tf")
    d["Scheduling Constraints"]=st.text_area(lbl("Scheduling constraints"),placeholder="e.g., 'Must avoid Monday mornings,' 'Cannot conflict with Academic Council'",key="m_con")
    d["Recurring"]=st.selectbox(lbl("Recurring?"),["No — one-time","Yes — weekly","Yes — biweekly","Yes — monthly","Yes — other cadence"],key="m_rec")
    d["Room Needed"]=st.radio(lbl("Room reservation needed?"),["Yes — please book","No — virtual or already booked","Not sure"],key="m_rm",horizontal=True)
    if d["Room Needed"]=="Yes — please book": d["Room Preferences"]=st.text_input("Room preferences:",key="m_rmd")
    d["Invitation Contents"]=", ".join(st.multiselect(lbl("Include in invitation"),["Agenda (provide or I'll draft)","Pre-read materials","Zoom/Teams link","Parking/location for external guests","Dial-in info"],key="m_inv"))
    d["Additional Support"]=", ".join(st.multiselect("Additional support needed",["Briefing materials","Agenda drafting","Note-taking or minutes","Follow-up communications","None"],key="m_sup"))
    d["Additional Notes"]=st.text_area("Additional context:",key="m_nt"); return d

def br_travel():
    d={}; sec("✈️ Travel & Logistics Coordination")
    d["Traveler(s)"]=st.text_input(lbl("Who is traveling?"),key="t_who")
    d["Travel Purpose"]=st.text_input(lbl("Purpose of travel"),placeholder="e.g., 'AACU Annual Conference'",key="t_pur")
    d["Travel Dates"]=st.text_input(lbl("Travel dates"),placeholder="Departure and return",key="t_dt")
    d["Destination"]=st.text_input(lbl("Destination"),placeholder="City, State, Country",key="t_dest")
    d["Arrangements Needed"]=", ".join(st.multiselect(lbl("Arrangements needed"),["Flight booking","Ground transportation","Hotel/lodging","Conference registration","Meal per diem setup","Parking","Visa/travel documents","Travel insurance"],key="t_arr"))
    d["Travel Preferences"]=st.text_area(lbl("Travel preferences or requirements"),key="t_prf")
    d["Funding Source"]=st.text_input(lbl("Funding source"),placeholder="Account/fund number or grant",key="t_fnd")
    d["Budget Cap"]=st.text_input("Budget cap:",key="t_cap")
    d["Travel Auth"]=st.radio(lbl("Travel authorization needed?"),["Yes — please initiate","Yes — already submitted","Not sure","No"],key="t_auth",horizontal=True)
    d["Briefing Materials"]=st.radio("Pre/post-travel briefing needed?",["Yes — itinerary packet","Yes — background research","No","Not sure"],key="t_brf",horizontal=True)
    d["Additional Notes"]=st.text_area("Additional notes:",key="t_nt"); return d

def br_document():
    d={}; sec("📝 Document Preparation or Editing")
    d["Document Title"]=st.text_input(lbl("Document name or working title"),key="d_tit")
    d["Document Type"]=st.selectbox(lbl("Type of document"),["— Select —","Report or summary","Policy or procedure","Agenda","Meeting minutes","Letter or correspondence","Memo","Contract or agreement","Proposal or grant narrative","Handbook or guide","Form or template","Spreadsheet or data table","Presentation slides","Other"],key="d_typ")
    d["Task"]=", ".join(st.multiselect(lbl("What do you need done?"),["Create from scratch","Edit or proofread","Reformat or redesign","Convert to different format","Combine/merge documents","Create a template"],key="d_tsk"))
    d["Intended Audience"]=st.text_input(lbl("Intended audience"),key="d_aud")
    d["Tone"]=st.radio(lbl("Desired tone"),["Formal and professional","Warm and collegial","Informational/neutral","Persuasive","Other"],key="d_ton",horizontal=True)
    d["Content / Outline"]=st.text_area(lbl("Content, outline, or key points"),key="d_cnt")
    d["Formatting Requirements"]=st.text_area(lbl("Formatting requirements"),key="d_fmt")
    d["Signatures / Approvals"]=st.radio("Signatures or approvals required?",["Yes — flag when ready","No","Not sure"],key="d_sig",horizontal=True)
    d["Additional Instructions"]=st.text_area("Additional instructions:",key="d_nt"); return d

def br_comm():
    d={}; sec("📣 Communication Drafting")
    d["Communication Type"]=st.selectbox(lbl("Type of communication"),["— Select —","Email (routine)","Email (sensitive or high-stakes)","Formal letter (on letterhead)","All-faculty announcement","All-staff announcement","Campus-wide announcement","Newsletter or bulletin item","Social media post","Press release or media statement","Thank-you or recognition message","Condolence message","Invitation","Talking points or remarks for a speech","Other"],key="c_typ")
    d["Sender"]=st.radio(lbl("Who is sending this?"),["Provost (on their behalf)","Dean of Faculty Development (on their behalf)","Both (joint)","From my own name as EA","Other"],key="c_snd",horizontal=False)
    d["Recipient"]=st.text_input(lbl("Intended recipient or audience"),key="c_rec")
    d["Primary Message"]=st.text_area(lbl("Primary message or purpose"),key="c_msg")
    d["Tone"]=st.radio(lbl("Desired tone"),["Warm and collegial","Formal and professional","Celebratory","Urgent but calm","Empathetic / sensitive","Informational","Persuasive"],key="c_ton",horizontal=False)
    d["Must Include"]=st.text_area(lbl("Must include"),placeholder="Quotes, dates, links, exact wording.",key="c_inc")
    d["Avoid"]=st.text_area("Avoid (optional):",key="c_avd")
    d["Target Send Date"]=st.text_input(lbl("When should this be sent?"),key="c_wen")
    d["Needs Approval"]=st.radio(lbl("Needs review before sending?"),["Yes — route back to me","Yes — route to another person","No — send once confirmed","Not sure"],key="c_apr",horizontal=False)
    if "another person" in str(d.get("Needs Approval","")): d["Approver"]=st.text_input("Who should approve?",key="c_aprd")
    d["Handle Distribution"]=st.radio(lbl("Should I also handle sending?"),["Yes — please send once approved","No — I will send it","Not sure"],key="c_dis",horizontal=True)
    d["Additional Context"]=st.text_area("Additional context:",key="c_nt"); return d

def br_room():
    d={}; sec("🏛️ Room / Space Reservation")
    d["Reservation Purpose"]=st.text_input(lbl("What is the reservation for?"),key="r_pur")
    d["Date(s) Needed"]=st.text_input(lbl("Date(s) needed"),key="r_dt")
    d["Time Needed"]=st.text_input(lbl("Time needed"),placeholder="Start, end, and setup/breakdown time",key="r_tm")
    d["Number of People"]=st.number_input(lbl("How many people?"),min_value=1,max_value=1000,value=10,key="r_ct")
    d["Room Preference"]=st.text_input(lbl("Preferred building or room"),key="r_prf")
    d["Setup Needed"]=", ".join(st.multiselect(lbl("Setup or equipment needed"),["Projector or display screen","Microphone / PA system","Video conferencing (Teams/Zoom)","Whiteboard or flip chart","Specific seating arrangement","Accessible setup","No special setup"],key="r_set"))
    d["Catering in Room"]=st.radio("Catering needed?",["Yes — separate request coming","Yes — please handle as well","No"],key="r_cat",horizontal=True)
    d["Recurring"]=st.radio(lbl("Recurring?"),["No — one-time","Yes — weekly","Yes — monthly","Yes — other"],key="r_rec",horizontal=True)
    d["Additional Notes"]=st.text_area("Additional notes:",key="r_nt"); return d

def br_catering():
    d={}; sec("🍽️ Catering & Hospitality")
    d["Catering For"]=st.text_input(lbl("What is the catering for?"),key="ca_for")
    d["Date & Time"]=st.text_input(lbl("Date and time needed"),placeholder="Date, arrival time, event end time",key="ca_dt")
    d["Delivery Location"]=st.text_input(lbl("Delivery / setup location"),key="ca_loc")
    d["Number Served"]=st.number_input(lbl("How many people?"),min_value=1,max_value=2000,value=15,key="ca_ct")
    d["Service Type"]=st.selectbox(lbl("Type of service"),["— Select —","Light refreshments","Breakfast service","Lunch service","Dinner or reception","Reception / heavy appetizers","Dessert or celebration cake","Coffee/beverage only","Full plated meal","Buffet style","Box lunches","Other"],key="ca_typ")
    d["Dietary Restrictions"]=st.text_area(lbl("Dietary restrictions or needs"),key="ca_dt2")
    d["Budget"]=st.text_input(lbl("Budget or budget range"),key="ca_bud")
    d["Account / Fund"]=st.text_input(lbl("Account or fund to charge"),key="ca_acc")
    d["Preferred Vendor"]=st.radio(lbl("Preferred vendor?"),["Use Berea College Dining Services","Specific external vendor","No preference — please recommend"],key="ca_ven",horizontal=False)
    if "external" in str(d.get("Preferred Vendor","")): d["Vendor Detail"]=st.text_input("Vendor name/contact:",key="ca_vend")
    d["Serving Supplies"]=st.radio(lbl("Serving supplies needed?"),["Yes — please coordinate","No — provided by venue","Not sure"],key="ca_sup",horizontal=True)
    d["Additional Notes"]=st.text_area("Additional notes:",key="ca_nt"); return d

def br_guest():
    d={}; sec("👤 Guest or Visitor Coordination")
    d["Guest(s) Info"]=st.text_area(lbl("Who is the guest / visitor?"),placeholder="Full name(s), title(s), institution, relationship to Berea",key="g_who")
    d["Visit Purpose"]=st.text_area(lbl("Purpose of this visit"),key="g_pur")
    d["Visit Dates & Times"]=st.text_input(lbl("Visit dates and times"),key="g_dt")
    d["Coordination Needed"]=", ".join(st.multiselect(lbl("Coordination needed"),["Create visit itinerary","Arrange transportation/parking","Book hotel","Arrange meals","Schedule meetings","Prepare welcome packet","Process honorarium","Send pre-visit info to guest","Coordinate with Security or Facilities","Coordinate A/V or tech support"],key="g_nds"))
    d["Berea POC"]=st.text_input(lbl("Berea College primary point of contact"),key="g_poc")
    d["Guest Contact"]=st.text_area(lbl("Guest's contact information"),key="g_cnt")
    d["Special Needs"]=st.text_area("Special needs, preferences, or sensitivities (optional):",key="g_spc")
    d["Hospitality Budget"]=st.text_input("Budget for hospitality (if applicable):",key="g_bud")
    d["Additional Notes"]=st.text_area("Additional notes:",key="g_nt"); return d

def br_budget():
    d={}; sec("💰 Budget, Purchasing & Expense Processing")
    d["Task Type"]=st.selectbox(lbl("Type of financial task"),["— Select —","Purchase order or vendor payment","Reimbursement processing","Budget tracking or reporting","Account reconciliation","Check request","Contract or service agreement routing","Grant-related expenditure","Other"],key="b_typ")
    d["Description"]=st.text_area(lbl("Describe what needs to be processed"),key="b_des")
    d["Account / Fund"]=st.text_input(lbl("Account or fund number"),key="b_acc")
    d["Amount"]=st.text_input(lbl("Total or estimated amount"),key="b_amt")
    d["Approval Status"]=st.radio(lbl("Has this been approved?"),["Yes — approved","Pending approval","Not yet — please advise","Unsure"],key="b_apr",horizontal=False)
    d["Vendor / Invoice Info"]=st.text_area("Vendor contact or invoice info:",key="b_ven")
    d["Payment Deadline"]=st.text_input(lbl("Payment deadline, if any"),key="b_ddl")
    d["Additional Instructions"]=st.text_area("Additional instructions:",key="b_nt"); return d

def br_research():
    d={}; sec("🔍 Research & Data Compilation")
    d["Research Topic"]=st.text_area(lbl("Topic or question to be researched"),key="re_top")
    d["Research Type"]=", ".join(st.multiselect(lbl("Type of research needed"),["Background info on a person or institution","Best practices or peer comparisons","Policy or regulatory research","Internal data (enrollment, faculty data)","Literature review","News or media coverage","Survey or assessment data","Budget or financial data","Other"],key="re_typ"))
    d["Sources"]=st.text_area("Sources to prioritize or avoid (optional):",key="re_src")
    d["Delivery Format"]=st.radio(lbl("Preferred delivery format"),["Brief summary (1 page)","Detailed report with citations","Bullet point briefing","Spreadsheet or data table","Presentation slides","Whatever makes most sense"],key="re_fmt",horizontal=False)
    d["How It Will Be Used"]=st.text_input(lbl("How will this research be used?"),key="re_use")
    d["Confidentiality"]=st.radio(lbl("Confidentiality concern?"),["Yes — handle with discretion","No — standard handling"],key="re_con",horizontal=True)
    d["Additional Context"]=st.text_area("Additional context:",key="re_nt"); return d

def br_facdev():
    d={}; sec("🎓 Faculty Development Programming Support")
    d["Initiative Name / Topic"]=st.text_input(lbl("Name or topic of the initiative"),key="fd_nam")
    d["Support Needed"]=", ".join(st.multiselect(lbl("Support needed"),["Program planning and logistics","Facilitator or presenter coordination","Participant communication and registration","Materials preparation","Room and technology setup","Pre/post-event survey","Certificate or recognition coordination","Assessment or outcomes tracking","Budget coordination","Reporting or program documentation","Website or portal update","Other"],key="fd_sup"))
    d["Timeline"]=st.text_area(lbl("Timeline"),placeholder="Key milestones, program date(s), and reporting deadlines.",key="fd_tim")
    d["Target Audience"]=", ".join(st.multiselect(lbl("Target audience"),["All full-time faculty","New faculty","Adjunct or part-time faculty","Department chairs and program directors","Academic staff","Graduate students","Specific department(s)","Other"],key="fd_aud"))
    d["Tied To"]=st.radio("Tied to grant, accreditation, or strategic initiative?",["Yes (describe below)","No","Not sure"],key="fd_tie",horizontal=True)
    if d.get("Tied To","")=="Yes (describe below)": d["Connection Detail"]=st.text_input("Describe the connection:",key="fd_tied")
    d["Key Collaborators"]=st.text_area("Key collaborators (optional):",key="fd_col")
    d["Additional Notes"]=st.text_area("Additional notes:",key="fd_nt"); return d

def br_pres():
    d={}; sec("📊 Presentation & Materials Preparation")
    d["Presentation Title"]=st.text_input(lbl("Title or topic"),key="p_tit")
    d["Presenter(s)"]=st.text_input(lbl("Who will be presenting?"),key="p_pre")
    d["Audience"]=st.text_input(lbl("Audience"),key="p_aud")
    d["Date & Location"]=st.text_input(lbl("When and where?"),key="p_whn")
    d["Materials Needed"]=", ".join(st.multiselect(lbl("Materials needed"),["Slides (create from scratch)","Slides (edit existing)","Talking points or speaker notes","Printed handouts","Data visualizations or charts","Executive summary or one-pager","Bios for speakers","Briefing book or pre-reading packet","Video or media to embed","Other"],key="p_mat"))
    d["Length / Format"]=st.text_input(lbl("Desired length or format"),key="p_fmt")
    d["Content / Key Messages"]=st.text_area(lbl("Content, data, or key messages"),key="p_cnt")
    d["Branding"]=st.radio(lbl("Branding guidelines?"),["Yes — use Berea College official templates","Yes — use a specific template (I will share it)","No — use your best judgment"],key="p_brd",horizontal=False)
    d["A/V Support"]=st.radio("A/V or tech support needed?",["Yes — please coordinate","No — already handled","Not sure"],key="p_av",horizontal=True)
    d["Additional Context"]=st.text_area("Additional context:",key="p_nt"); return d

def br_other():
    d={}; sec("🗂️ Other / General Administrative Support")
    d["Support Description"]=st.text_area(lbl("Describe the support needed in detail"),key="o_des")
    d["Key Steps / Deliverables"]=st.text_area(lbl("Key steps or deliverables, as far as you know"),key="o_stp")
    d["People / Contacts Involved"]=st.text_area("People or departments involved (optional):",key="o_ppl")
    d["Templates / Resources"]=st.text_area("Templates or resources to reference (optional):",key="o_res")
    d["Sensitivities"]=st.text_area(lbl("Time-sensitive or politically sensitive considerations?"),key="o_sen")
    return d

BRANCHES={
    "📅  Event Planning & Coordination":br_event,
    "🗓️  Meeting Scheduling & Calendar Management":br_meeting,
    "✈️  Travel & Logistics Coordination":br_travel,
    "📝  Document Preparation or Editing":br_document,
    "📣  Communication Drafting (Emails, Letters, Memos, Announcements)":br_comm,
    "🏛️  Room / Space Reservation Only":br_room,
    "🍽️  Catering & Hospitality":br_catering,
    "👤  Guest or Visitor Coordination":br_guest,
    "💰  Budget, Purchasing & Expense Processing":br_budget,
    "🔍  Research & Data Compilation":br_research,
    "🎓  Faculty Development Programming Support":br_facdev,
    "📊  Presentation & Materials Preparation":br_pres,
    "🗂️  Other / General Administrative Support":br_other,
}

# ── Shared: Notes thread ─────────────────────────────────────────────────────
def notes_thread(ref,complete=False):
    notes=DB.get_notes(ref)
    sec("💬 Notes thread")
    if not notes: st.caption("No notes yet.")
    for n in notes:
        is_t=n.get("author_role","")=="admin"
        cls="tiff-note" if is_t else ""
        mc="tiff" if is_t else ""
        who="Tiff (Provost's Office)" if is_t else n.get("author","")
        st.markdown(f'<div class="note-card {cls}"><div class="note-meta {mc}">{who} &nbsp;·&nbsp; {n.get("timestamp","")}</div><div>{n.get("text","")}</div></div>',unsafe_allow_html=True)
    if complete:
        st.info("This request is complete. Submit a new request if you need further assistance.")
        return
    with st.form(key=f"nf_{ref}",clear_on_submit=True):
        txt=st.text_area("Add a note:",placeholder="e.g., 'Date changed to Sept 5' or 'Please also include the registrar.'",key=f"ni_{ref}")
        if st.form_submit_button("Send note →"):
            if txt.strip():
                author=cur_user(); role="admin" if is_admin() else "supervisor"
                DB.add_note(ref,author,role,txt.strip())
                subs=DB.get_submissions()
                sub=next((s for s in subs if s.get("ref_number")==ref),{})
                if is_admin():
                    sub_email=sub.get("Notification Email","") or sub.get("Contact Info","")
                    notify_new_note(ref,"Tiff (Provost's Office)",txt.strip(),sub.get("Your Name",""),sub_email)
                else:
                    notify_new_note(ref,author,txt.strip(),"__admin__",ADMIN_EMAIL)
                st.success("Note sent — the other party has been notified by email.")
                st.rerun()
            else: st.warning("Please enter a note before sending.")

# ── View: Login ──────────────────────────────────────────────────────────────
def v_login():
    banner("Administrative Support Request System")
    _,col,_=st.columns([1,2,1])
    with col:
        st.markdown(f'<div style="background:{WHITE};border:1px solid #cdd6e0;border-top:5px solid {BLUE};border-radius:10px;padding:1.8rem 1.6rem;text-align:center;margin-top:.8rem;"><div style="font-size:2rem;color:{BLUE};margin-bottom:7px;">🔒</div><h3 style="font-family:\'Newsreader\',serif;color:{BLUE};font-size:1.15rem;font-weight:600;margin-bottom:5px;">Secure access</h3><p style="font-size:.84rem;color:#555;margin-bottom:1rem;">For authorized Berea College personnel only.</p></div>',unsafe_allow_html=True)
        st.markdown("<br>",unsafe_allow_html=True)
        with st.form("login_form",clear_on_submit=False):
            email_in=st.text_input("Email address",placeholder="you@berea.edu",key="li_em")
            password=st.text_input("Password",type="password",placeholder="Enter your password",key="li_pw")
            submitted=st.form_submit_button("Sign in →")
        if submitted:
            em=email_in.strip().lower()
            # Admin: matches the admin email (if set) OR just the admin password as a fallback.
            admin_email=(ADMIN_EMAIL or "").strip().lower()
            if password==ADMIN_PASSWORD and (not admin_email or em==admin_email or em==""):
                st.session_state.user="__admin__"; st.session_state.user_role="admin"
                st.session_state.view="admin"; st.rerun()
            else:
                matched=False
                for name,info in SUPERVISORS.items():
                    sup_email=str(info.get("email","")).strip().lower()
                    if em==sup_email and password==info.get("password",""):
                        st.session_state.user=name; st.session_state.user_role="supervisor"
                        st.session_state.view="portal"; matched=True; st.rerun(); break
                if not matched: st.error("Incorrect email or password. Please try again or contact the Provost's office.")
        st.markdown(f'<p style="text-align:center;font-size:.73rem;color:#aaa;margin-top:.8rem;">🔐 Encrypted connection (HTTPS) · Authorized personnel only</p>',unsafe_allow_html=True)

# ── View: Form ────────────────────────────────────────────────────────────────
def v_form():
    if st.session_state.submitted: v_confirm(); return
    banner("Submit a Support Request")
    st.markdown('<div class="ro-notice">🔐 Encrypted · Fields marked <strong>*</strong> are required</div>',unsafe_allow_html=True)
    sec("👤 About This Request")
    # Build the list of supervisor names from the configured accounts.
    sup_names=sorted(SUPERVISORS.keys())
    name_options=["— Select your name —"]+sup_names+["Other (type below)"]
    # If a supervisor is logged in, default the dropdown to their own name.
    default_idx=0
    if is_sup() and cur_user() in sup_names:
        default_idx=name_options.index(cur_user())
    c1,c2=st.columns(2)
    with c1:
        name_choice=st.selectbox(lbl("Your name"),name_options,index=default_idx,key="u_nm")
    with c2:
        # Role is pulled automatically from the chosen supervisor; shown read-only.
        auto_role=SUPERVISORS.get(name_choice,{}).get("role","") if name_choice in sup_names else ""
        st.text_input("Your role / title",value=auto_role,disabled=True,key="u_role_display",
                      help="This fills in automatically based on the name you select.")
    # Resolve the actual name + role values used on submit.
    if name_choice=="Other (type below)":
        name=st.text_input(lbl("Type your name"),placeholder="First and last name",key="u_nm_other")
        role=st.text_input("Your role / title",placeholder="e.g., Associate Provost",key="u_role_other")
    elif name_choice in sup_names:
        name=name_choice
        role=auto_role
    else:
        name=""
        role=""
    # Pre-fill contact with the supervisor's email on file (they can edit it).
    default_contact=SUPERVISORS.get(name,{}).get("email","") if name in sup_names else ""
    contact=st.text_input(lbl("Preferred contact for follow-up"),value=default_contact,
                          placeholder="Email and/or Teams handle",key="u_con")
    hr()
    c3,c4=st.columns(2)
    with c3: deadline=st.date_input(lbl("Deadline or ideal completion date"),min_value=date.today(),key="u_ddl")
    with c4: urgency=st.selectbox(lbl("Urgency"),["— Select —","🔴 Urgent — needed within 24–48 hours","🟠 High Priority — needed within the week","🟡 Standard — needed within 2 weeks","🟢 Low — no hard deadline"],key="u_urg")
    conf=st.radio(lbl("Is this request confidential?"),["Yes — please handle with discretion","No — standard handling"],key="u_conf",horizontal=True)
    if "Yes" in conf: st.info("🔒 Marked confidential.")
    hr(); sec("🗂️ Request Type")
    rtype=st.selectbox(lbl("What type of support are you requesting?"),REQUEST_TYPES,key="u_rt")
    bd={}
    if rtype in BRANCHES: hr(); bd=BRANCHES[rtype]()
    hr(); sec("📎 Final Details")
    uploaded=st.file_uploader("Attachments (optional)",accept_multiple_files=True,key="u_files")
    files=", ".join([f.name for f in uploaded]) if uploaded else ""
    final=st.text_area("Anything else I should know?",placeholder="Context, history, sensitivities, dependencies.",key="u_final")
    st.markdown("**Notification preferences for this request:**")
    n_stat=st.checkbox("Email me when the status of this request changes",value=True,key="pf_st")
    n_note=st.checkbox("Email me when a note is added to this request",value=True,key="pf_nt")
    n_email=st.text_input("Email for notifications (if different from contact above):",key="pf_em")
    hr()
    if st.button("✅  Submit Support Request",key="sub_btn"):
        errs=[]
        if not name.strip(): errs.append("Please select or enter your name.")
        if not contact.strip(): errs.append("Preferred contact is required.")
        if urgency=="— Select —": errs.append("Urgency is required.")
        if rtype=="— Select a request type —": errs.append("Please select a request type.")
        if errs:
            for e in errs: st.error(f"⚠️ {e}")
            return
        ref=gen_ref()
        data={"Your Name":name,"Your Role":role,"Contact Info":contact,
              "Deadline":deadline.strftime("%Y-%m-%d"),"Urgency":urgency,
              "Confidential":conf,"Request Type":rtype,"Status":"New"}
        data.update(bd)
        data["File Attachments"]=files; data["Additional Notes"]=final
        data["Notify on Status Change"]=str(n_stat); data["Notify on New Note"]=str(n_note)
        data["Notification Email"]=n_email or contact
        data["submitted_at"]=datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        DB.save_submission(data,ref); notify_new_submission(data,ref)
        DB.save_prefs(name,{"on_new_note":n_note,"on_status_change":n_stat,"on_new_submission":False,"on_any_note":False})
        st.session_state.submitted=True; st.session_state.submission_data=data
        st.session_state.ref_number=ref; st.rerun()

def v_confirm():
    data=st.session_state.submission_data; ref=st.session_state.ref_number
    banner("✅ Request Submitted")
    uc=ucss(data.get("Urgency",""))
    st.markdown(f'<div style="background:#f0f5fb;border:2px solid {BLUE};border-top:5px solid {CHART};border-radius:10px;padding:1.4rem 1.4rem 1.1rem;margin-top:.5rem;"><h3 style="font-family:\'Newsreader\',serif;color:{BLUE};font-size:1.15rem;font-weight:600;margin:0 0 5px;">Thank you, {data.get("Your Name","")}!</h3><p style="font-size:.84rem;color:#555;margin-bottom:9px;">Your request has been received. Tiff will follow up shortly.</p><p style="margin-bottom:3px;font-size:.84rem;"><strong>Reference:</strong><br><span style="background:{BLUE};color:{CHART};font-family:\'Barlow Condensed\',monospace;font-size:.95rem;font-weight:700;letter-spacing:.12em;padding:3px 11px;border-radius:4px;display:inline-block;margin:3px 0 7px;">{ref}</span></p><p style="font-size:.84rem;margin:0;"><strong>Type:</strong> {data.get("Request Type","")}<br><strong>Urgency:</strong> <span class="sb {uc}">{data.get("Urgency","")}</span> &nbsp; <strong>Deadline:</strong> {data.get("Deadline","")}</p></div>',unsafe_allow_html=True)
    sm=summary(data,ref); st.markdown("<br>",unsafe_allow_html=True)
    with st.expander("📋 View full submission"): st.text(sm)
    st.download_button("⬇️  Download my submission",data=sm,file_name=f"{ref}.txt",mime="text/plain")
    st.markdown("<br>",unsafe_allow_html=True)
    if st.button("📝  Submit another request",key="anb"):
        st.session_state.submitted=False; st.session_state.submission_data={}
        st.session_state.ref_number=""; st.rerun()

# ── Shared accept / deny actions ─────────────────────────────────────────────
def accept_request(sub):
    """Accept a request: mark Accepted, create a linked task, notify submitter."""
    ref=sub.get("ref_number","")
    name=sub.get("Your Name","")
    rt=sub.get("Request Type","").split("  ")[-1] if "  " in sub.get("Request Type","") else sub.get("Request Type","")
    urg=sub.get("Urgency","")
    DB.update_status(ref,"Accepted")
    DB.save_task({"task_id":gen_task_id(),
        "created_at":datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "Title":f"{rt} — {name}","Category":"Follow-up / admin",
        "Priority":("🔴 Urgent" if "Urgent" in urg else "🟠 High" if "High" in urg else "🟡 Medium"),
        "Deadline":sub.get("Deadline",""),"Status":"In Progress","Recurrence":"None — one-time",
        "Details":f"Accepted from request {ref} ({name}, {sub.get('Your Role','')}).\n\n{sub.get('Additional Notes','')}","source_ref":ref,"Visibility":"Shared"})
    se=sub.get("Notification Email","") or sub.get("Contact Info","")
    notify_accepted(ref,rt,name,se)

def deny_request(sub,reason=""):
    """Deny a request: mark Declined, notify submitter (no task created)."""
    ref=sub.get("ref_number","")
    name=sub.get("Your Name","")
    rt=sub.get("Request Type","").split("  ")[-1] if "  " in sub.get("Request Type","") else sub.get("Request Type","")
    DB.update_status(ref,"Declined")
    se=sub.get("Notification Email","") or sub.get("Contact Info","")
    notify_declined(ref,rt,name,se,reason)

# ── View: Admin dashboard ────────────────────────────────────────────────────
def v_admin():
    banner("Admin Dashboard — Everything on My Plate","Berea College · Office of the Provost · Admin view")
    subs=DB.get_submissions()
    tasks=DB.get_tasks()
    # Refs that have been accepted/converted into tasks — these now count ONLY as tasks
    converted_refs={t.get("source_ref","") for t in tasks if t.get("source_ref","")}
    active_tasks=[t for t in tasks if t.get("Status","")!="Complete"]
    # Requests awaiting acceptance: not complete, not declined, not yet converted.
    pending_reqs=[s for s in subs if s.get("Status","") not in ["Complete","Declined"]
                  and s.get("ref_number","") not in converted_refs]
    # Urgent/high across both active tasks and pending requests.
    urg_ct=(sum(1 for s in pending_reqs if "Urgent" in s.get("Urgency","") or "High" in s.get("Urgency",""))
            +sum(1 for t in active_tasks if "Urgent" in t.get("Priority","") or "High" in t.get("Priority","")))
    done_ct=(sum(1 for s in subs if s.get("Status","")=="Complete")
             +sum(1 for t in tasks if t.get("Status","")=="Complete"))

    def _card(col,num,label,cs,btn_label,target_view):
        with col:
            st.markdown(f'<div class="mc"><div class="mn {cs}">{num}</div><div class="ml">{label}</div></div>',unsafe_allow_html=True)
            if st.button(btn_label,key=f"card_{target_view}",use_container_width=True):
                st.session_state.view=target_view; st.rerun()

    c1,c2,c3,c4=st.columns(4)
    _card(c1,len(active_tasks),"On my plate (active tasks)","bl","Go to my tasks →","tasks")
    _card(c2,len(pending_reqs),"Pending acceptance","bl","Review & respond →","pending")
    _card(c3,urg_ct,"Urgent / high","am","View urgent →","urgent")
    _card(c4,done_ct,"Completed","gr","View completed →","completed")
    st.caption("Tip: the buttons under each number jump you straight to that list.")

    # ── Gantt timeline (accepted work, filterable by supervisor) ────────────
    st.markdown("<br>",unsafe_allow_html=True)
    sec("📊 Timeline (Gantt view)")
    sup_names=sorted(SUPERVISORS.keys())
    filter_opts=["All supervisors"]+sup_names+["My own tasks"]
    gcol1,gcol2=st.columns([2,3])
    with gcol1:
        gfilter=st.selectbox("Show requests from",filter_opts,key="gantt_filter")
    st.caption("Accepted requests and tasks, drawn from acceptance date to deadline. "
               "Completed items are greyed out. Use the dropdown to focus on one supervisor or see everyone at once.")
    draw_gantt(gantt_rows_admin(gfilter),legend_label="SUPERVISOR")

    st.markdown("<br>",unsafe_allow_html=True)
    ca,cb,cc,cd=st.columns([2,1.5,1.5,1])
    with ca: srch=st.text_input("Search",placeholder="Name, reference, or request type…",key="ad_s",label_visibility="collapsed")
    with cb: fu=st.selectbox("Urgency",["All urgency levels","🔴 Urgent","🟠 High Priority","🟡 Standard","🟢 Low"],key="ad_u",label_visibility="collapsed")
    with cc: fs=st.selectbox("Status",["All statuses"]+STATUS_OPTIONS,key="ad_fs",label_visibility="collapsed")
    with cd:
        if subs:
            rows=[{k:(json.dumps(v) if isinstance(v,list) else str(v)) for k,v in s.items()} for s in subs]
            st.download_button("⬇️ CSV",data=pd.DataFrame(rows).to_csv(index=False),
                               file_name=f"requests_{datetime.now().strftime('%Y%m%d')}.csv",mime="text/csv",key="ad_dl")
    filt=subs
    if srch: s2=srch.lower(); filt=[x for x in filt if s2 in x.get("Your Name","").lower() or s2 in x.get("ref_number","").lower() or s2 in x.get("Request Type","").lower()]
    if fu!="All urgency levels": filt=[x for x in filt if fu.split(" ",1)[-1] in x.get("Urgency","")]
    if fs!="All statuses": filt=[x for x in filt if x.get("Status","")==fs]
    st.caption(f"Showing {len(filt)} of {len(subs)} requests")
    if not filt: st.info("No requests match your filters."); return
    for sub in filt:
        ref=sub.get("ref_number",""); name=sub.get("Your Name","")
        rt=sub.get("Request Type","").split("  ")[-1] if "  " in sub.get("Request Type","") else sub.get("Request Type","")
        urg=sub.get("Urgency",""); stat=sub.get("Status","New")
        sc=S_CSS.get(stat,"s-new"); uc=ucss(urg)
        conf="Yes" in str(sub.get("Confidential",""))
        cf='<span class="conf-flag">🔒 Conf.</span>' if conf else ""
        short=ref.split("-")[-1] if ref else "—"
        _conv=any(t.get("source_ref","")==ref for t in tasks)
        _acc="✅ " if _conv else ""
        with st.expander(f"{_acc}{'🔴 ' if 'Urgent' in urg else '🟠 ' if 'High' in urg else ''}{short}  ·  {name}  ·  {rt}  ·  {stat}"):
            cl,cr=st.columns([2,1])
            with cl:
                st.markdown(f'<div style="font-size:.82rem;"><span class="ref-code">{ref}</span> &nbsp; {cf}<br><strong>{name}</strong> · {sub.get("Your Role","")}<br><strong>Type:</strong> {sub.get("Request Type","")}<br><strong>Urgency:</strong> <span class="sb {uc}">{urg}</span> &nbsp; <strong>Deadline:</strong> {sub.get("Deadline","")}<br><strong>Submitted:</strong> {sub.get("submitted_at","")}</div>',unsafe_allow_html=True)
                dj=sub.get("Details JSON",""); ex={}
                if dj:
                    try: ex=json.loads(dj)
                    except: pass
                af={k:v for k,v in {**sub,**ex}.items() if k not in ["ref_number","submitted_at","Status","Details JSON","Your Name","Your Role","Contact Info","Deadline","Urgency","Confidential","Request Type"]}
                if af:
                    with st.expander("📋 Full submission details",expanded=False):
                        for k,v in af.items():
                            if v and str(v) not in ["","—","False"]: st.markdown(f"**{k}:** {v}")
            with cr:
                # Has this request already been converted to a task?
                existing_tasks=DB.get_tasks()
                already_converted=any(t.get("source_ref","")==ref for t in existing_tasks)

                if not already_converted and stat not in ["Complete","Declined"]:
                    st.markdown("**Respond to this request:**")
                    st.caption("Accept creates a linked task and notifies the submitter. Deny notifies them courteously.")
                    if st.button("✅ Accept & convert to task",key=f"acc_{ref}"):
                        accept_request(sub)
                        st.success("Accepted — task created and submitter notified.")
                        st.rerun()
                    if st.button("🚫 Deny this request",key=f"deny_{ref}"):
                        deny_request(sub)
                        st.warning("Request declined — submitter notified.")
                        st.rerun()
                    st.markdown("<br>",unsafe_allow_html=True)

                if already_converted:
                    st.markdown(f'<div class="ro-notice">✅ Accepted — this request is now tracked as a task in <strong>My tasks</strong>.</div>',unsafe_allow_html=True)

                st.markdown("**Update status:**")
                ns=st.selectbox("Status:",STATUS_OPTIONS,index=STATUS_OPTIONS.index(stat) if stat in STATUS_OPTIONS else 0,key=f"st_{ref}",label_visibility="collapsed")
                if st.button("Save status",key=f"sv_{ref}"):
                    if ns!=stat:
                        DB.update_status(ref,ns)
                        se=sub.get("Notification Email","") or sub.get("Contact Info","")
                        notify_status_change(ref,ns,name,se)
                        st.success(f"Updated to: {ns}"); st.rerun()
            hr(); notes_thread(ref,complete=(stat=="Complete"))

# ── View: Pending acceptance (accept / deny queue) ───────────────────────────
def view_pending():
    banner("Pending Acceptance","Berea College · Office of the Provost · Requests awaiting your response")
    subs=DB.get_submissions()
    tasks=DB.get_tasks()
    converted_refs={t.get("source_ref","") for t in tasks if t.get("source_ref","")}
    pending=[s for s in subs if s.get("Status","") not in ["Complete","Declined"]
             and s.get("ref_number","") not in converted_refs]
    if st.button("← Back to dashboard",key="pend_back"):
        st.session_state.view="admin"; st.rerun()
    st.caption(f"{len(pending)} request{'s' if len(pending)!=1 else ''} waiting for you to accept or deny. "
               "Accepting creates a linked task; denying sends a courteous note to the submitter.")
    if not pending:
        st.success("🎉 Nothing waiting — you're all caught up on new requests.")
        return
    for sub in sorted(pending,key=lambda x:x.get("submitted_at","")):
        ref=sub.get("ref_number",""); name=sub.get("Your Name","")
        rt=sub.get("Request Type","").split("  ")[-1] if "  " in sub.get("Request Type","") else sub.get("Request Type","")
        urg=sub.get("Urgency",""); uc=ucss(urg)
        conf="Yes" in str(sub.get("Confidential",""))
        cf='<span class="conf-flag">🔒 Conf.</span>' if conf else ""
        with st.expander(f"{'🔴 ' if 'Urgent' in urg else '🟠 ' if 'High' in urg else ''}{name}  ·  {rt}  ·  {sub.get('Deadline','')}"):
            st.markdown(f'<div style="font-size:.82rem;"><span class="ref-code">{ref}</span> &nbsp; {cf}<br><strong>{name}</strong> · {sub.get("Your Role","")}<br><strong>Type:</strong> {sub.get("Request Type","")}<br><strong>Urgency:</strong> <span class="sb {uc}">{urg}</span> &nbsp; <strong>Deadline:</strong> {sub.get("Deadline","")}<br><strong>Submitted:</strong> {sub.get("submitted_at","")}</div>',unsafe_allow_html=True)
            dj=sub.get("Details JSON",""); ex={}
            if dj:
                try: ex=json.loads(dj)
                except: pass
            af={k:v for k,v in {**sub,**ex}.items() if k not in ["ref_number","submitted_at","Status","Details JSON","Your Name","Your Role","Contact Info","Deadline","Urgency","Confidential","Request Type"]}
            if af:
                with st.expander("📋 Full submission details",expanded=False):
                    for k,v in af.items():
                        if v and str(v) not in ["","—","False"]: st.markdown(f"**{k}:** {v}")
            hr()
            b1,b2=st.columns(2)
            with b1:
                if st.button("✅ Accept",key=f"pacc_{ref}",use_container_width=True):
                    accept_request(sub); st.success("Accepted — task created and submitter notified."); st.rerun()
            with b2:
                if st.button("🚫 Deny",key=f"pden_{ref}",use_container_width=True):
                    deny_request(sub); st.warning("Declined — submitter notified."); st.rerun()

# ── View: Urgent / high priority items ───────────────────────────────────────
def view_urgent():
    banner("Urgent & High Priority","Berea College · Office of the Provost · Everything needing attention first")
    subs=DB.get_submissions()
    tasks=DB.get_tasks()
    converted_refs={t.get("source_ref","") for t in tasks if t.get("source_ref","")}
    if st.button("← Back to dashboard",key="urg_back"):
        st.session_state.view="admin"; st.rerun()
    # Urgent/high active tasks.
    u_tasks=[t for t in tasks if t.get("Status","")!="Complete"
             and ("Urgent" in t.get("Priority","") or "High" in t.get("Priority",""))]
    # Urgent/high pending requests (not yet accepted).
    u_reqs=[s for s in subs if s.get("Status","") not in ["Complete","Declined"]
            and s.get("ref_number","") not in converted_refs
            and ("Urgent" in s.get("Urgency","") or "High" in s.get("Urgency",""))]
    st.caption(f"{len(u_tasks)} urgent/high task{'s' if len(u_tasks)!=1 else ''} and "
               f"{len(u_reqs)} urgent/high request{'s' if len(u_reqs)!=1 else ''} awaiting acceptance.")
    if not u_tasks and not u_reqs:
        st.success("Nothing urgent right now.")
        return
    if u_tasks:
        sec("🔴 Urgent / high tasks")
        for t in sorted(u_tasks,key=lambda x:x.get("Deadline","")):
            pri=t.get("Priority",""); puc=("urg-u" if "Urgent" in pri else "urg-h")
            st.markdown(f'<div class="note-card"><strong>{t.get("Title","")}</strong> '
                        f'&nbsp; <span class="sb {puc}">{pri}</span> '
                        f'&nbsp; <span class="sb {S_CSS.get(t.get("Status","New"),"s-new")}">{t.get("Status","New")}</span><br>'
                        f'<span style="font-size:.78rem;color:#5a7a96;">{t.get("Category","")} · due {t.get("Deadline","")}</span></div>',
                        unsafe_allow_html=True)
    if u_reqs:
        sec("🟠 Urgent / high requests (awaiting acceptance)")
        for s in sorted(u_reqs,key=lambda x:x.get("Deadline","")):
            urg=s.get("Urgency",""); uc=ucss(urg)
            rt=s.get("Request Type","").split("  ")[-1] if "  " in s.get("Request Type","") else s.get("Request Type","")
            st.markdown(f'<div class="note-card"><strong>{s.get("Your Name","")}</strong> — {rt} '
                        f'&nbsp; <span class="sb {uc}">{urg}</span><br>'
                        f'<span style="font-size:.78rem;color:#5a7a96;">due {s.get("Deadline","")}</span></div>',
                        unsafe_allow_html=True)
        if st.button("Go to Pending Acceptance to respond →",key="urg_to_pending"):
            st.session_state.view="pending"; st.rerun()

# ── View: Supervisor portal ───────────────────────────────────────────────────
def v_portal():
    user=cur_user(); info=SUPERVISORS.get(user,{}); role_t=info.get("role","")
    banner(f"My Support Requests",f"Berea College · {role_t} · {user}")
    subs=DB.get_submissions(name=user)
    total=len(subs)
    # Each request falls into exactly ONE bucket below (no overlap), so the
    # four cards always add up to the total number of requests.
    pending=sum(1 for s in subs if s.get("Status","") in ["New","Under Review",""])
    prog=sum(1 for s in subs if s.get("Status","") in ["Accepted","In Progress","On Hold"])
    awt=sum(1 for s in subs if s.get("Status","")=="Awaiting Info")
    done=sum(1 for s in subs if s.get("Status","")=="Complete")
    c1,c2,c3,c4=st.columns(4)
    for col,n,lb,cs in [(c1,pending,"Submitted & pending confirmation of availability","bl"),
                        (c2,prog,"In progress","am"),
                        (c3,awt,"Needs your input","am"),
                        (c4,done,"Complete","gr")]:
        col.markdown(f'<div class="mc"><div class="mn {cs}">{n}</div><div class="ml">{lb}</div></div>',unsafe_allow_html=True)
    st.caption(f"You have submitted **{total}** request{'s' if total!=1 else ''} in total. Each one is counted in exactly one card above based on its current status — so you can always see where things stand at a glance.")
    st.markdown(f'<div class="ro-notice" style="margin-top:.6rem;">👁️ Requests are <strong>read-only</strong> once submitted. To flag a change or update, add a note — Tiff will be notified by email.</div>',unsafe_allow_html=True)
    if not subs: st.info("No requests yet. Use the sidebar to submit your first request."); return
    # ── Timeline of everything this supervisor has submitted ────────────────
    st.markdown("<br>",unsafe_allow_html=True)
    sec("📊 My Timeline (Gantt view)")
    admin_opts=["All",ADMIN_NAME]  # forward-looking: more admins can be added here later
    hcol1,hcol2=st.columns([2,3])
    with hcol1:
        handled_by=st.selectbox("Handled by",admin_opts,key="sup_gantt_handler")
    st.caption("Every request you've submitted, drawn to its deadline. Colored by request type; "
               "completed items are greyed out. (Currently all requests are handled by "
               f"{ADMIN_NAME}; the filter will expand if more administrators are added.)")
    # There's only one admin today, so the filter shows everything regardless of choice.
    draw_gantt(gantt_rows_supervisor(user),legend_label="TYPE")
    # Explain what each status means for the submitter.
    st.markdown(f'<div style="margin:.8rem 0 .4rem;font-size:.78rem;color:#5a7a96;line-height:1.7;">'
                f'<strong style="font-family:\'Barlow Condensed\',sans-serif;letter-spacing:.04em;text-transform:uppercase;font-size:.7rem;">What each status means:</strong><br>'
                f'<span class="sb s-new">New</span> submitted, pending confirmation of availability &nbsp; '
                f'<span class="sb s-review">Under Review</span> Tiff is reviewing it &nbsp; '
                f'<span class="sb s-prog">Accepted</span> confirmed &amp; underway &nbsp; '
                f'<span class="sb s-prog">In Progress</span> being worked on &nbsp; '
                f'<span class="sb s-await">Awaiting Info</span> needs something from you &nbsp; '
                f'<span class="sb s-done">Complete</span> finished &nbsp; '
                f'<span class="sb s-hold">On Hold</span> paused'
                f'</div>',unsafe_allow_html=True)
    st.markdown("<br>",unsafe_allow_html=True)
    for sub in sorted(subs,key=lambda x:x.get("submitted_at",""),reverse=True):
        ref=sub.get("ref_number",""); rt=sub.get("Request Type","").split("  ")[-1] if "  " in sub.get("Request Type","") else sub.get("Request Type","")
        urg=sub.get("Urgency",""); stat=sub.get("Status","New")
        sc=S_CSS.get(stat,"s-new"); uc=ucss(urg)
        nct=len(DB.get_notes(ref)); short=ref.split("-")[-1] if ref else "—"
        fstat=friendly_status(stat)
        with st.expander(f"{short}  ·  {rt}  ·  {fstat}  ·  {f'{nct} note(s)' if nct else 'No notes'}"):
            st.markdown(f'<div style="background:{WHITE};border:0.5px solid #cdd6e0;border-left:4px solid {BLUE};border-radius:0 8px 8px 0;padding:11px 13px;margin-bottom:9px;font-size:.82rem;"><div style="font-size:.7rem;color:#5a7a96;margin-bottom:6px;">🔒 Read-only · Submitted {sub.get("submitted_at","")} · <span class="ref-code">{ref}</span></div><div><strong>Type:</strong> {sub.get("Request Type","")}</div><div><strong>Urgency:</strong> <span class="sb {uc}">{urg}</span> &nbsp; <strong>Deadline:</strong> {sub.get("Deadline","")}</div><div style="margin-top:5px;"><strong>Status:</strong> <span class="sb {sc}">{fstat}</span></div></div>',unsafe_allow_html=True)
            dj=sub.get("Details JSON",""); ex={}
            if dj:
                try: ex=json.loads(dj)
                except: pass
            af={k:v for k,v in {**sub,**ex}.items() if k not in ["ref_number","submitted_at","Status","Details JSON"]}
            with st.expander("📋 View full request details",expanded=False):
                for k,v in af.items():
                    if v and str(v) not in ["","—","False"]: st.markdown(f"**{k}:** {v}")
            hr(); notes_thread(ref,complete=(stat=="Complete"))

    # ── What Tiff is working on (shared tasks) ──────────────────────────────
    all_tasks=DB.get_tasks()
    shared=[t for t in all_tasks if str(t.get("Visibility","Shared"))!="Private"]
    shared_active=[t for t in shared if t.get("Status","")!="Complete"]
    st.markdown("<br>",unsafe_allow_html=True)
    sec("🗂️ What Tiff is working on")
    st.caption(f"For transparency, here's everything currently on Tiff's plate — **{len(shared_active)} active item{'s' if len(shared_active)!=1 else ''}**. "
               f"These are read-only, but you're welcome to add a comment on any of them; Tiff will be notified.")
    if not shared_active:
        st.info("Nothing shared on Tiff's task list right now.")
    else:
        for tk in sorted(shared_active,key=lambda x:x.get("Deadline","")):
            tid=tk.get("task_id",""); ttl=tk.get("Title","(untitled)")
            tpri=tk.get("Priority",""); tstat=tk.get("Status","New")
            tsc=S_CSS.get(tstat,"s-new"); tuc=ucss(tpri)
            tcat=tk.get("Category",""); tddl=tk.get("Deadline","")
            tnc=len(DB.get_notes(tid))
            src=tk.get("source_ref","")
            src_note=f" · from request {src.split('-')[-1]}" if src else ""
            with st.expander(f"{tpri.split(' ')[0] if tpri else ''} {ttl}  ·  {tstat}  ·  {f'{tnc} comment(s)' if tnc else 'No comments'}"):
                det=tk.get("Details","")
                det_html=f'<div style="margin-top:6px;color:#333;">{det}</div>' if det else ''
                st.markdown(f'<div style="background:{WHITE};border:0.5px solid #cdd6e0;border-left:4px solid {BLUE};border-radius:0 8px 8px 0;padding:11px 13px;margin-bottom:9px;font-size:.82rem;">'
                            f'<div style="font-size:.7rem;color:#5a7a96;margin-bottom:6px;">🔒 Read-only{src_note}</div>'
                            f'<div><strong>Category:</strong> {tcat}</div>'
                            f'<div><strong>Priority:</strong> <span class="sb {tuc}">{tpri}</span> &nbsp; <strong>Deadline:</strong> {tddl}</div>'
                            f'<div style="margin-top:5px;"><strong>Status:</strong> <span class="sb {tsc}">{tstat}</span></div>'
                            f'{det_html}'
                            f'</div>',unsafe_allow_html=True)
                task_notes_thread(tid,complete=(tstat=="Complete"),task_title=ttl,allow_comment=True)

# ── Shared: task notes thread (admin-only, on personal tasks) ─────────────────
def task_notes_thread(task_id,complete=False,task_title="",allow_comment=True):
    notes=DB.get_notes(task_id)
    sec("💬 Comments")
    if not notes: st.caption("No comments yet.")
    for n in notes:
        is_admin_note=n.get("author_role","")=="admin"
        cls="tiff-note" if is_admin_note else ""
        who=(n.get("author","") or "Tiff") if is_admin_note else n.get("author","")
        st.markdown(f'<div class="note-card {cls}"><div class="note-meta {"tiff" if is_admin_note else ""}">{who} &nbsp;·&nbsp; {n.get("timestamp","")}</div><div>{n.get("text","")}</div></div>',unsafe_allow_html=True)
    if complete or not allow_comment:
        return
    with st.form(key=f"tnf_{task_id}",clear_on_submit=True):
        txt=st.text_area("Add a comment:",placeholder="e.g., 'Any update on this?' or 'Happy to help move this along.'",key=f"tni_{task_id}")
        if st.form_submit_button("Add comment →"):
            if txt.strip():
                author=cur_user() or "Tiff"
                role="admin" if is_admin() else "supervisor"
                DB.add_note(task_id,author,role,txt.strip())
                # If a supervisor commented, notify Tiff.
                if role=="supervisor":
                    notify_task_comment(task_id,task_title or "(task)",author,txt.strip())
                    st.success("Comment added — Tiff has been notified.")
                else:
                    st.success("Comment added.")
                st.rerun()
            else: st.warning("Please enter a comment.")

# ── View: My Tasks (admin only — Tiff's own to-dos) ──────────────────────────
def draw_gantt(rows, legend_label="WHO"):
    """Draw an SVG timeline from prepared rows. Each row is a dict with:
    title, start (datetime), end (datetime), status (str), group (color key), complete (bool).
    Bars grey out when complete. Colored by 'group'."""
    from datetime import datetime as _dt, timedelta as _td2
    if not rows:
        st.info("Nothing to show on the timeline yet.")
        return
    palette=["#004175","#49B4EF","#7D901E","#C75B12","#8AAAC5","#3B6D11","#9b1c1c","#005A8B"]
    groups=[]
    for r in rows:
        if r["group"] not in groups: groups.append(r["group"])
    color_of={g:palette[i%len(palette)] for i,g in enumerate(groups)}
    min_d=min(r["start"] for r in rows); max_d=max(r["end"] for r in rows)
    span=max((max_d-min_d).days,1); pad=max(int(span*0.05),2)
    min_d=min_d-_td2(days=pad); max_d=max_d+_td2(days=pad)
    total_days=max((max_d-min_d).days,1)
    rows=sorted(rows,key=lambda r:(r["start"],r["end"]))
    row_h=30; gap=8; top=54; left=250; right=30
    width=980; plot_w=width-left-right
    height=top+len(rows)*(row_h+gap)+30
    def x_of(d): return left+plot_w*((d-min_d).days/total_days)
    parts=[f'<svg viewBox="0 0 {width} {height}" xmlns="http://www.w3.org/2000/svg" style="width:100%;height:auto;font-family:Barlow,sans-serif;">']
    parts.append(f'<rect x="0" y="0" width="{width}" height="{height}" fill="#ffffff"/>')
    cur=_dt(min_d.year,min_d.month,1)
    while cur<max_d:
        gx=x_of(cur)
        if gx>=left:
            parts.append(f'<line x1="{gx:.1f}" y1="{top-8}" x2="{gx:.1f}" y2="{height-20}" stroke="#e3e9f0" stroke-width="1"/>')
            parts.append(f'<text x="{gx+4:.1f}" y="{top-14}" font-size="11" fill="#5a7a96" font-weight="600">{cur.strftime("%b %Y")}</text>')
        cur=_dt(cur.year+1,1,1) if cur.month==12 else _dt(cur.year,cur.month+1,1)
    today=_dt.now()
    if min_d<=today<=max_d:
        tx=x_of(today)
        parts.append(f'<line x1="{tx:.1f}" y1="{top-8}" x2="{tx:.1f}" y2="{height-20}" stroke="#C75B12" stroke-width="2" stroke-dasharray="4 3"/>')
        parts.append(f'<text x="{tx+4:.1f}" y="{height-8}" font-size="10" fill="#C75B12" font-weight="700">Today</text>')
    for i,r in enumerate(rows):
        y=top+i*(row_h+gap)
        bx=x_of(r["start"]); bw=max(x_of(r["end"])-bx,6)
        fill="#c9ced6" if r["complete"] else color_of[r["group"]]
        opacity="0.55" if r["complete"] else "1"
        label=r["title"] if len(r["title"])<=34 else r["title"][:32]+"…"
        deco=' text-decoration="line-through"' if r["complete"] else ''
        parts.append(f'<text x="10" y="{y+row_h*0.66:.1f}" font-size="12" fill="#1a1a1a"{deco}>{label}</text>')
        parts.append(f'<rect x="{bx:.1f}" y="{y:.1f}" width="{bw:.1f}" height="{row_h}" rx="5" fill="{fill}" opacity="{opacity}"/>')
        tag=friendly_status(r["status"]) if not r["complete"] else "Complete"
        short_tag={"Submitted & pending confirmation of availability":"Pending","Confirmed & underway":"Accepted","In progress":"In progress","Awaiting your input":"Awaiting info","Under review by Tiff":"Under review","On hold":"On hold","Complete":"Complete"}.get(tag,tag)
        if bw>70:
            parts.append(f'<text x="{bx+8:.1f}" y="{y+row_h*0.66:.1f}" font-size="10.5" fill="#ffffff" font-weight="600">{short_tag}</text>')
    ly=height-6; lx=left
    parts.append(f'<text x="10" y="{ly:.1f}" font-size="10.5" fill="#5a7a96" font-weight="700">{legend_label}:</text>')
    for g in groups:
        parts.append(f'<rect x="{lx:.1f}" y="{ly-9:.1f}" width="11" height="11" rx="2" fill="{color_of[g]}"/>')
        gl=g if len(g)<=22 else g[:20]+"…"
        parts.append(f'<text x="{lx+15:.1f}" y="{ly:.1f}" font-size="10.5" fill="#333">{gl}</text>')
        lx+=25+len(gl)*6.2
    parts.append('</svg>')
    st.markdown("".join(parts),unsafe_allow_html=True)

def _parse_date(s):
    from datetime import datetime as _dt
    s=str(s or "").strip()
    try: return _dt.strptime(s[:10],"%Y-%m-%d")
    except: return None

def gantt_rows_admin(supervisor_filter="All supervisors"):
    """Admin view: accepted work only (tasks), from acceptance date → deadline,
    grouped/colored by supervisor. Optional filter to one supervisor or own tasks."""
    tasks=DB.get_tasks()
    subs=DB.get_submissions()
    sub_by_ref={s.get("ref_number",""):s for s in subs}
    rows=[]
    for t in tasks:
        start=_parse_date(t.get("created_at","")); end=_parse_date(t.get("Deadline",""))
        if not start or not end: continue
        if end<start: end=start
        src=t.get("source_ref","")
        who=sub_by_ref[src].get("Your Name","") if src in sub_by_ref else "My own tasks"
        who=who or "My own tasks"
        if supervisor_filter!="All supervisors" and who!=supervisor_filter:
            continue
        rows.append({"title":t.get("Title","(untitled)"),"start":start,"end":end,
                     "status":t.get("Status","New"),"group":who,
                     "complete":t.get("Status","")=="Complete"})
    return rows

def gantt_rows_supervisor(name):
    """Supervisor view: ALL requests they submitted, from acceptance date (if accepted)
    else submission date → deadline, grouped/colored by request type."""
    subs=DB.get_submissions(name=name)
    tasks=DB.get_tasks()
    task_by_ref={t.get("source_ref",""):t for t in tasks if t.get("source_ref","")}
    rows=[]
    for s in subs:
        ref=s.get("ref_number","")
        # Start: acceptance date if a linked task exists, else submission date.
        start=None
        if ref in task_by_ref:
            start=_parse_date(task_by_ref[ref].get("created_at",""))
        if not start:
            start=_parse_date(s.get("submitted_at",""))
        end=_parse_date(s.get("Deadline",""))
        if not start or not end: continue
        if end<start: end=start
        rtype=s.get("Request Type","")
        rtype_short=rtype.split("  ")[-1] if "  " in rtype else rtype
        rows.append({"title":s.get("Request Type","").split("  ")[-1] if "  " in s.get("Request Type","") else s.get("Request Type","(request)"),
                     "start":start,"end":end,"status":s.get("Status","New"),
                     "group":rtype_short or "Other","complete":s.get("Status","")=="Complete"})
    return rows

def v_tasks():
    banner("My Tasks","Berea College · Office of the Provost · Tiff's personal task list")
    tasks=DB.get_tasks()
    active=[t for t in tasks if t.get("Status","")!="Complete"]
    done=[t for t in tasks if t.get("Status","")=="Complete"]
    urgent=sum(1 for t in active if "Urgent" in t.get("Priority","") or "High" in t.get("Priority",""))
    recurring=sum(1 for t in tasks if t.get("Recurrence","") not in ["","None — one-time"])
    c1,c2,c3,c4=st.columns(4)
    for col,n,lb,cs in [(c1,len(active),"Active tasks","bl"),(c2,urgent,"Urgent / high","am"),
                        (c3,recurring,"Recurring","bl"),(c4,len(done),"Completed","gr")]:
        col.markdown(f'<div class="mc"><div class="mn {cs}">{n}</div><div class="ml">{lb}</div></div>',unsafe_allow_html=True)

    st.markdown("<br>",unsafe_allow_html=True)
    sec("➕ Add a New Task")
    with st.form("new_task",clear_on_submit=True):
        tc1,tc2=st.columns([2,1])
        with tc1: t_title=st.text_input(lbl("Task title"),placeholder="e.g., 'Draft April APC minutes'")
        with tc2: t_cat=st.selectbox(lbl("Category"),TASK_CATEGORIES)
        tc3,tc4,tc5=st.columns(3)
        with tc3: t_pri=st.selectbox(lbl("Priority"),TASK_PRIORITIES,index=2)
        with tc4: t_ddl=st.date_input(lbl("Deadline"),min_value=date.today())
        with tc5: t_rec=st.selectbox("Recurrence",RECURRENCE_OPTIONS)
        t_det=st.text_area("Details / notes",placeholder="Any specifics, links, or context for this task.")
        t_share=st.checkbox("Share this task with supervisors (visible on their dashboard)",value=True,
                            help="On by default. Uncheck to keep this task private to you.")
        if st.form_submit_button("➕ Add task"):
            if not t_title.strip():
                st.error("⚠️ Task title is required.")
            else:
                tid=gen_task_id()
                DB.save_task({"task_id":tid,"created_at":datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "Title":t_title.strip(),"Category":t_cat,"Priority":t_pri,
                    "Deadline":t_ddl.strftime("%Y-%m-%d"),"Status":"New","Recurrence":t_rec,
                    "Details":t_det.strip(),"source_ref":"","Visibility":"Shared" if t_share else "Private"})
                st.success(f"Task added: {t_title.strip()}"); st.rerun()

    if not tasks:
        st.info("No tasks yet. Add your first task above."); return

    st.markdown("<br>",unsafe_allow_html=True)
    sec("📋 Active Tasks")
    fcol1,fcol2=st.columns([1.5,1.5])
    with fcol1: fcat=st.selectbox("Filter by category",["All categories"]+TASK_CATEGORIES,key="tf_cat",label_visibility="collapsed")
    with fcol2: fpri=st.selectbox("Filter by priority",["All priorities"]+TASK_PRIORITIES,key="tf_pri",label_visibility="collapsed")

    shown=active
    if fcat!="All categories": shown=[t for t in shown if t.get("Category","")==fcat]
    if fpri!="All priorities": shown=[t for t in shown if t.get("Priority","")==fpri]

    if not shown: st.caption("No active tasks match your filters.")
    for t in sorted(shown,key=lambda x:x.get("Deadline","")):
        tid=t.get("task_id",""); pri=t.get("Priority","")
        uc=("urg-u" if "Urgent" in pri else "urg-h" if "High" in pri else "urg-s" if "Medium" in pri else "urg-l")
        stat=t.get("Status","New"); sc=S_CSS.get(stat,"s-new")
        rec=t.get("Recurrence","")
        rec_badge=f' &nbsp; <span class="sb s-review">🔁 {rec}</span>' if rec not in ["","None — one-time"] else ""
        src=t.get("source_ref","")
        src_badge=f' &nbsp; <span class="sb s-hold">from {src.split("-")[-1]}</span>' if src else ""
        vis=str(t.get("Visibility","Shared"))
        vis_badge=(' &nbsp; <span class="sb s-done">👁️ Shared with supervisors</span>' if vis!="Private"
                   else ' &nbsp; <span class="sb s-hold">🔒 Private</span>')
        with st.expander(f"{pri.split(' ')[0]} {t.get('Title','')}  ·  {t.get('Category','')}  ·  {stat}"):
            st.markdown(f'<div style="font-size:.82rem;"><strong>Category:</strong> {t.get("Category","")}<br><strong>Priority:</strong> <span class="sb {uc}">{pri}</span> &nbsp; <strong>Deadline:</strong> {t.get("Deadline","")}{rec_badge}{src_badge}{vis_badge}<br><strong>Created:</strong> {t.get("created_at","")}</div>',unsafe_allow_html=True)
            if t.get("Details",""): st.markdown(f'<div style="font-size:.84rem;margin-top:6px;background:#f5f8fc;border-radius:6px;padding:8px 11px;">{t.get("Details","")}</div>',unsafe_allow_html=True)
            hr()
            uc1,uc2,uc3=st.columns(3)
            with uc1:
                ns=st.selectbox("Status",STATUS_OPTIONS,index=STATUS_OPTIONS.index(stat) if stat in STATUS_OPTIONS else 0,key=f"ts_{tid}")
            with uc2:
                np=st.selectbox("Priority",TASK_PRIORITIES,index=TASK_PRIORITIES.index(pri) if pri in TASK_PRIORITIES else 2,key=f"tp_{tid}")
            with uc3:
                nv=st.selectbox("Visibility",["Shared","Private"],index=0 if vis!="Private" else 1,key=f"tv_{tid}")
            if st.button("💾 Update",key=f"tu_{tid}"):
                if ns!=stat:
                    DB.update_task(tid,"Status",ns)
                    # Recurring task completed → spawn next occurrence
                    if ns=="Complete" and rec not in ["","None — one-time"]:
                        nd=next_occurrence(t.get("Deadline",""),rec)
                        DB.save_task({"task_id":gen_task_id(),"created_at":datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "Title":t.get("Title",""),"Category":t.get("Category",""),
                            "Priority":pri,"Deadline":nd,"Status":"New","Recurrence":rec,
                            "Details":t.get("Details",""),"source_ref":"","Visibility":t.get("Visibility","Shared")})
                        st.success(f"Completed. Next occurrence created for {nd}.")
                if np!=pri: DB.update_task(tid,"Priority",np)
                if nv!=vis: DB.update_task(tid,"Visibility",nv)
                st.rerun()
            with st.expander("🗑️ Delete this task"):
                st.caption("This permanently removes the task.")
                if st.button("Confirm delete",key=f"td_{tid}"):
                    DB.delete_task(tid); st.rerun()
            hr(); task_notes_thread(tid,complete=(stat=="Complete"))

    if done:
        st.markdown("<br>",unsafe_allow_html=True)
        with st.expander(f"✅ Completed tasks ({len(done)})"):
            for t in sorted(done,key=lambda x:x.get("Deadline",""),reverse=True):
                st.markdown(f"• ~~{t.get('Title','')}~~ · {t.get('Category','')} · completed")

# ── View: Notification preferences ───────────────────────────────────────────
def v_prefs():
    user=cur_user(); banner("Notification Preferences")
    info=SUPERVISORS.get(user,{}) if is_sup() else {}
    em=info.get("email","—") if is_sup() else (ADMIN_EMAIL or "Not configured")
    st.markdown(f'<div class="ro-notice">📧 Notifications sent to: <strong>{em}</strong></div>',unsafe_allow_html=True)
    prefs=DB.get_prefs(user); sec("🔔 Email Notification Settings")
    with st.form("pf_form"):
        if is_sup():
            p1=st.checkbox("Email me when the **status** of one of my requests changes",value=prefs.get("on_status_change",True))
            p2=st.checkbox("Email me when **Tiff adds a note** to one of my requests",value=prefs.get("on_new_note",True))
            new_p={"on_status_change":p1,"on_new_note":p2,"on_new_submission":False,"on_any_note":False}
        else:
            p1=st.checkbox("Email me when a **new request** is submitted",value=prefs.get("on_new_submission",True))
            p2=st.checkbox("Email me when a supervisor **adds a note** to any request",value=prefs.get("on_any_note",True))
            p3=st.checkbox("Email me when I **update a status** (confirmation copy)",value=prefs.get("on_status_change",False))
            new_p={"on_new_submission":p1,"on_any_note":p2,"on_status_change":p3,"on_new_note":False}
        st.markdown("<br>",unsafe_allow_html=True)
        if st.form_submit_button("Save preferences"):
            DB.save_prefs(user,new_p); st.success("Preferences saved.")
    hr(); sec("📬 About Notifications")
    st.markdown("All notifications are sent from the configured Gmail account. They are never sent to external parties and never include confidential request content — only the reference number and a brief description. To update your email address, contact Tiff to update the app's Secrets configuration.")

# ── Main router ───────────────────────────────────────────────────────────────
def main():
    if not st.session_state.user: v_login(); return
    sidebar(); v=st.session_state.get("view","form")
    if v=="admin" and is_admin(): v_admin()
    elif v=="tasks" and is_admin(): v_tasks()
    elif v=="pending" and is_admin(): view_pending()
    elif v=="urgent" and is_admin(): view_urgent()
    elif v=="completed" and is_admin(): v_tasks()
    elif v=="portal" and is_sup(): v_portal()
    elif v=="prefs": v_prefs()
    else: v_form()

if __name__=="__main__": main()
