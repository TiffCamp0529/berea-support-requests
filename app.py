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
        # Cache verified worksheets so header sync only runs once per sheet per session.
        if not hasattr(self,"_verified_ws"): self._verified_ws=set()
        try:
            ws=self._wb.worksheet(name)
            if hdrs and name not in self._verified_ws:
                self._verified_ws.add(name)
                try:
                    existing=ws.row_values(1) if ws.row_count>0 else []
                except: existing=[]
                missing=[h for h in hdrs if h not in existing]
                if missing:
                    try:
                        # Add all missing headers in one batch.
                        start_col=len(existing)+1
                        for i,h in enumerate(missing):
                            ws.update_cell(1,start_col+i,h)
                    except: pass  # If header update fails, reads may still work.
            return ws
        except:
            ws=self._wb.add_worksheet(title=name,rows=2000,cols=max(len(hdrs),10))
            if hdrs: ws.append_row(hdrs)
            self._verified_ws.add(name)
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
                ws=self._ws("submissions",self.SH)
                all_vals=ws.get_all_values()
                if len(all_vals)<2: return []
                headers=all_vals[0]
                rows=[]
                for row in all_vals[1:]:
                    if not any(str(c).strip() for c in row): continue
                    record={headers[i]:(row[i] if i<len(row) else "") for i in range(len(headers))}
                    rows.append(record)
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
                all_vals=ws.get_all_values()
                if len(all_vals)<2: return []
                headers=all_vals[0]
                results=[]
                for row in all_vals[1:]:
                    if not any(str(c).strip() for c in row): continue
                    record={headers[i]:(row[i] if i<len(row) else "") for i in range(len(headers))}
                    if record.get("ref_number","")==ref: results.append(record)
                return results
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
        "Status","Recurrence","Details","source_ref","Visibility","Checklist",
        "Calendar_Start","Calendar_End","Owner","AssignedTo","StaffVisible"]
    def save_task(self,task):
        if self.ok:
            ws=self._ws("tasks",self.TH); ws.append_row([str(task.get(h,"")) for h in self.TH])
        else:
            if "tasks" not in st.session_state: st.session_state.tasks=[]
            st.session_state.tasks.append(task)
    def get_tasks(self):
        if self.ok:
            try:
                ws=self._ws("tasks",self.TH)
                # Use get_all_values + manual dict building for reliability
                # (get_all_records can fail when rows are shorter than headers).
                all_vals=ws.get_all_values()
                if len(all_vals)<2: return []
                headers=all_vals[0]
                tasks=[]
                for row in all_vals[1:]:
                    if not any(str(c).strip() for c in row): continue  # skip empty rows
                    record={headers[i]:(row[i] if i<len(row) else "") for i in range(len(headers))}
                    tasks.append(record)
                return tasks
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
    # ── DELEGATIONS (staff-to-staff task requests) ────────────────────────────
    DGH=["deleg_id","created_at","from_user","to_email","to_name","Title","Category",
         "Priority","Deadline","Details","Status","task_id"]
    def save_delegation(self,d):
        if self.ok:
            ws=self._ws("delegations",self.DGH); ws.append_row([str(d.get(h,"")) for h in self.DGH])
        else:
            if "delegations" not in st.session_state: st.session_state.delegations=[]
            st.session_state.delegations.append(d)
    def get_delegations(self,to_email=None,status=None):
        rows=[]
        if self.ok:
            try:
                ws=self._ws("delegations",self.DGH); all_vals=ws.get_all_values()
                if len(all_vals)>=2:
                    headers=all_vals[0]
                    for row in all_vals[1:]:
                        if any(str(c).strip() for c in row):
                            rows.append({headers[i]:(row[i] if i<len(row) else "") for i in range(len(headers))})
            except: rows=[]
        else:
            rows=list(st.session_state.get("delegations",[]))
        if to_email: rows=[r for r in rows if str(r.get("to_email","")).strip().lower()==to_email.strip().lower()]
        if status: rows=[r for r in rows if r.get("Status","")==status]
        return rows
    def update_delegation(self,deleg_id,field,value):
        if self.ok:
            try:
                ws=self._ws("delegations",self.DGH); cell=ws.find(deleg_id)
                if cell:
                    h=ws.row_values(1); ws.update_cell(cell.row,h.index(field)+1,str(value))
            except: pass
        else:
            for d in st.session_state.get("delegations",[]):
                if d.get("deleg_id")==deleg_id: d[field]=value
    # ── USERS (stored in Google Sheets, hashed passwords) ─────────────────────
    UH=["user_id","email","display_name","role","title","password_hash","active","created_at"]
    def _hash_pw(self,pw):
        import hashlib; return hashlib.sha256(pw.encode()).hexdigest()
    def get_users(self,raw=False):
        """Get all users. Deduplicates by email, keeping the latest entry (last row wins)."""
        records=[]
        if self.ok:
            try:
                ws=self._ws("users",self.UH)
                all_vals=ws.get_all_values()
                if len(all_vals)>=2:
                    headers=all_vals[0]
                    for row in all_vals[1:]:
                        if any(str(c).strip() for c in row):
                            records.append({headers[i]:(row[i] if i<len(row) else "") for i in range(len(headers))})
            except: records=[]
        else:
            records=list(st.session_state.get("app_users",[]))
        if raw: return records
        # Deduplicate by email — last entry wins (most recently added/updated).
        seen={}
        for u in records:
            em=str(u.get("email","")).strip().lower()
            if em: seen[em]=u
        return list(seen.values())
    def get_user_by_email(self,email):
        for u in self.get_users():
            if str(u.get("email","")).strip().lower()==email.strip().lower() and str(u.get("active","True")).lower()!="false":
                return u
        return None
    def save_user(self,user):
        if self.ok:
            ws=self._ws("users",self.UH)
            ws.append_row([str(user.get(h,"")) for h in self.UH])
        else:
            if "app_users" not in st.session_state: st.session_state.app_users=[]
            st.session_state.app_users.append(user)
    def update_user(self,user_id,field,value):
        if self.ok:
            try:
                ws=self._ws("users",self.UH); cell=ws.find(user_id)
                if cell:
                    h=ws.row_values(1); ws.update_cell(cell.row,h.index(field)+1,str(value))
            except: pass
        else:
            for u in st.session_state.get("app_users",[]):
                if u.get("user_id")==user_id: u[field]=value
    def cleanup_duplicate_users(self):
        """Remove duplicate user rows from the sheet, keeping only the last entry per email."""
        if not self.ok: return
        try:
            ws=self._ws("users",self.UH)
            all_vals=ws.get_all_values()
            if len(all_vals)<2: return
            headers=all_vals[0]
            email_col=headers.index("email") if "email" in headers else -1
            if email_col<0: return
            # Walk rows from bottom to top, keep first occurrence of each email, delete the rest.
            seen=set(); rows_to_delete=[]
            for i in range(len(all_vals)-1,0,-1):  # skip header row
                em=all_vals[i][email_col].strip().lower() if email_col<len(all_vals[i]) else ""
                if not em: rows_to_delete.append(i+1); continue  # empty email = junk row
                if em in seen: rows_to_delete.append(i+1)  # duplicate
                else: seen.add(em)
            for row_num in sorted(rows_to_delete,reverse=True):
                ws.delete_rows(row_num)
        except: pass
    def seed_users_from_secrets(self):
        """Migrate admin + supervisor accounts from secrets into the users sheet.
        Only adds users whose email isn't already present. Runs cleanup first."""
        self.cleanup_duplicate_users()
        existing=self.get_users()
        existing_emails={str(u.get("email","")).strip().lower() for u in existing}
        import uuid
        ae=(ADMIN_EMAIL or "").strip().lower()
        if ae and ae not in existing_emails:
            self.save_user({"user_id":str(uuid.uuid4())[:8],"email":ae,
                "display_name":ADMIN_NAME,"role":"admin","title":"Administrator",
                "password_hash":self._hash_pw(ADMIN_PASSWORD),
                "active":"True","created_at":datetime.now().strftime("%Y-%m-%d %H:%M:%S")})
        for name,info in SUPERVISORS.items():
            se=str(info.get("email","")).strip().lower()
            if se and se not in existing_emails:
                self.save_user({"user_id":str(uuid.uuid4())[:8],"email":se,
                    "display_name":name,"role":"supervisor","title":info.get("role",""),
                    "password_hash":self._hash_pw(info.get("password","")),
                    "active":"True","created_at":datetime.now().strftime("%Y-%m-%d %H:%M:%S")})
    # ── AUDIT TRAIL ───────────────────────────────────────────────────────────
    AH=["timestamp","ref_or_id","action","old_value","new_value","by_user","by_role"]
    def log_audit(self,ref_or_id,action,old_val="",new_val="",by_user="",by_role=""):
        entry={"timestamp":datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
               "ref_or_id":ref_or_id,"action":action,
               "old_value":str(old_val),"new_value":str(new_val),
               "by_user":by_user,"by_role":by_role}
        if self.ok:
            try:
                ws=self._ws("audit_log",self.AH)
                ws.append_row([entry[h] for h in self.AH])
            except: pass
        else:
            if "audit_log" not in st.session_state: st.session_state.audit_log=[]
            st.session_state.audit_log.append(entry)
    def get_audit(self,ref_or_id=None):
        if self.ok:
            try:
                ws=self._ws("audit_log",self.AH); rows=ws.get_all_records()
                if ref_or_id: return [r for r in rows if r.get("ref_or_id")==ref_or_id]
                return rows
            except: return []
        logs=list(st.session_state.get("audit_log",[]))
        if ref_or_id: return [r for r in logs if r.get("ref_or_id")==ref_or_id]
        return logs

@st.cache_resource
def get_db(): return DataLayer()
DB=get_db()
# Seed users from secrets — only once per session.
if DB.ok and not st.session_state.get("_users_seeded",False):
    try:
        DB.seed_users_from_secrets()
        st.session_state._users_seeded=True
    except: pass

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

def is_admin(): return st.session_state.user_role in ("admin",)
def is_support(): return st.session_state.user_role in ("support",)
def is_colleague(): return st.session_state.user_role in ("colleague",)
def is_admin_or_support(): return st.session_state.user_role in ("admin","support","colleague")
def is_staff(): return st.session_state.user_role in ("admin","support","colleague")
def can_delegate(): return st.session_state.user_role in ("admin","colleague")
def is_sup(): return st.session_state.user_role=="supervisor"
def cur_user(): return st.session_state.user
def active_mode(): return st.session_state.get("active_mode",st.session_state.user_role)

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
STATUS_OPTIONS=["New","Under Review","Accepted","In Progress","Awaiting Info","Complete","On Hold","Declined","Cancelled"]
S_CSS={"New":"s-new","Under Review":"s-review","Accepted":"s-prog","In Progress":"s-prog",
       "Awaiting Info":"s-await","Complete":"s-done","On Hold":"s-hold","Declined":"s-await","Cancelled":"s-hold"}

def banner(t,s="Berea College · Office of the Provost & Dean of Faculty Development"):
    st.markdown(f'<div class="banner"><h1>{t}</h1><p>{s}</p></div>',unsafe_allow_html=True)
def sec(l): st.markdown(f'<div class="sec">{l}</div>',unsafe_allow_html=True)
def hr(): st.markdown('<hr class="hr">',unsafe_allow_html=True)
def pii_banner(compact=False):
    """Confidentiality / PII warning. Use compact=True for inline placement near input fields."""
    if compact:
        st.markdown('<div style="background:#FFF3CD;border-left:4px solid #C75B12;border-radius:0 6px 6px 0;'
                    'padding:7px 11px;margin:8px 0;font-size:.76rem;color:#5a3e00;line-height:1.5;">'
                    '🔒 <strong>Do not enter confidential information, PII (Social Security numbers, student records, '
                    'medical information, etc.), or sensitive documents here.</strong> This platform is not a secure '
                    'storage location for protected data. Instead, email sensitive materials directly or share a link '
                    'to a secure location such as <strong>Box</strong>.</div>',unsafe_allow_html=True)
    else:
        st.markdown('<div style="background:#FFF3CD;border:1px solid #f0d080;border-left:5px solid #C75B12;'
                    'border-radius:0 8px 8px 0;padding:12px 16px;margin:12px 0;font-size:.82rem;color:#5a3e00;line-height:1.6;">'
                    '🔒 <strong>Confidentiality Notice</strong><br>'
                    'This application is <strong>not a secure storage platform</strong> for personally identifiable '
                    'information (PII), student records (FERPA), medical information (HIPAA), Social Security numbers, '
                    'financial account details, or other protected data.<br><br>'
                    '<strong>Do not upload or type confidential information into any field on this platform.</strong> '
                    'If your request involves sensitive materials, please email them directly to the appropriate '
                    'party or share a link to a secure, institutionally approved location such as '
                    '<strong>Berea College Box</strong>.<br><br>'
                    '<em>You may reference that confidential materials exist (e.g., "see attached via Box") '
                    'without including the materials themselves.</em></div>',unsafe_allow_html=True)
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
        "Cancelled":"Cancelled by you",
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
        if is_admin_or_support():
            if st.button("📊  Admin dashboard",key="nb1",use_container_width=True):
                st.session_state.view="admin"; st.session_state.submitted=False; st.rerun()
            if st.button("✅  My tasks",key="nb1b",use_container_width=True):
                st.session_state.view="tasks"; st.session_state.submitted=False; st.rerun()
            if st.button("🎯  Today's focus",key="nb1c",use_container_width=True):
                st.session_state.view="today"; st.session_state.submitted=False; st.rerun()
            if st.button("📆  Weekly planner",key="nb1d",use_container_width=True):
                st.session_state.view="planner"; st.session_state.submitted=False; st.rerun()
            _my_email=st.session_state.get("user_email","")
            _pending_dg=len(DB.get_delegations(to_email=_my_email,status="Pending")) if _my_email else 0
            _dg_label=f"📥  Task requests ({_pending_dg})" if _pending_dg else "📥  Task requests"
            if st.button(_dg_label,key="nb1e",use_container_width=True):
                st.session_state.view="delegations"; st.session_state.submitted=False; st.rerun()
            if st.button("📝  Submit new request",key="nb2",use_container_width=True):
                st.session_state.view="form"; st.session_state.submitted=False; st.rerun()
            if is_admin():
                if st.button("👥  Manage users",key="nb_users",use_container_width=True):
                    st.session_state.view="users"; st.rerun()
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
        st.markdown('<div style="margin-top:8px;background:#FFF8E8;border-radius:6px;padding:6px 9px;'
                    'font-size:.65rem;color:#7a5a00;line-height:1.4;">'
                    '🔒 <strong>No PII or confidential data.</strong> '
                    'Do not upload protected information. Use Box or email for sensitive materials.</div>',
                    unsafe_allow_html=True)
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
    d["Location"]=st.selectbox(lbl("Location"),["— Select —","On campus — already identified","On campus — TBD, please assist","Off campus — venue identified","Off campus — venue TBD, please assist","Virtual (Teams, Zoom, or other)","Hybrid","Multiple locations"],key="e_loc")
    if d["Location"]=="Multiple locations":
        st.caption("List each location. Start with up to 3; click 'Add another' for more.")
        if "e_loc_count" not in st.session_state: st.session_state.e_loc_count=3
        locs=[]
        for i in range(st.session_state.e_loc_count):
            v=st.text_input(f"Location {i+1}",key=f"e_mloc_{i}",placeholder=f"e.g., room, building, or off-campus venue")
            if v.strip(): locs.append(v.strip())
        if st.button("➕ Add another location",key="e_loc_more"):
            st.session_state.e_loc_count+=1; st.rerun()
        d["Location Detail"]="; ".join(locs) if locs else ""
    else:
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
    d["Room Preference"]=st.selectbox(lbl("Room / location"),["Single location","Multiple locations"],key="r_loc_type")
    if d["Room Preference"]=="Multiple locations":
        st.caption("List each room or location needed.")
        if "r_loc_count" not in st.session_state: st.session_state.r_loc_count=3
        rlocs=[]
        for i in range(st.session_state.r_loc_count):
            v=st.text_input(f"Room/location {i+1}",key=f"r_mloc_{i}",placeholder="e.g., Draper 101")
            if v.strip(): rlocs.append(v.strip())
        if st.button("➕ Add another location",key="r_loc_more"):
            st.session_state.r_loc_count+=1; st.rerun()
        d["Room Preference"]="; ".join(rlocs) if rlocs else "Multiple — see details"
    else:
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
        pii_banner(compact=True)
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
            # Try users sheet first (the primary auth source once seeded).
            sheet_user=DB.get_user_by_email(em) if em else None
            if sheet_user and DB._hash_pw(password)==sheet_user.get("password_hash",""):
                role=sheet_user.get("role","supervisor")
                dname=sheet_user.get("display_name","")
                st.session_state.user=dname or em
                st.session_state.user_role=role
                st.session_state.user_email=em
                st.session_state.active_mode=role  # for admin/support toggle
                # Route by role.
                if role in ("admin","support","colleague"):
                    st.session_state.view="admin"
                else:
                    st.session_state.view="portal"
                st.rerun()
            else:
                # Fallback: try secrets-based auth (backwards compatibility during migration).
                admin_email=(ADMIN_EMAIL or "").strip().lower()
                if password==ADMIN_PASSWORD and (not admin_email or em==admin_email or em==""):
                    st.session_state.user=ADMIN_NAME; st.session_state.user_role="admin"
                    st.session_state.user_email=admin_email; st.session_state.active_mode="admin"
                    st.session_state.view="admin"; st.rerun()
                else:
                    matched=False
                    for name,info in SUPERVISORS.items():
                        sup_email=str(info.get("email","")).strip().lower()
                        if em==sup_email and password==info.get("password",""):
                            st.session_state.user=name; st.session_state.user_role="supervisor"
                            st.session_state.user_email=em; st.session_state.active_mode="supervisor"
                            st.session_state.view="portal"; matched=True; st.rerun(); break
                    if not matched: st.error("Incorrect email or password. Please try again or contact the Provost's office.")
        st.markdown(f'<p style="text-align:center;font-size:.73rem;color:#aaa;margin-top:.8rem;">🔐 Encrypted connection (HTTPS) · Authorized personnel only</p>',unsafe_allow_html=True)
        st.markdown('<p style="text-align:center;font-size:.68rem;color:#b08030;margin-top:.4rem;">'
                    '⚠️ This platform is not approved for confidential or personally identifiable information (PII). '
                    'Do not upload protected data.</p>',unsafe_allow_html=True)

# ── View: Form ────────────────────────────────────────────────────────────────
def v_form():
    if st.session_state.submitted: v_confirm(); return
    banner("Submit a Support Request")
    st.markdown('<div class="ro-notice">🔐 Encrypted · Fields marked <strong>*</strong> are required</div>',unsafe_allow_html=True)
    pii_banner()
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
    pii_banner(compact=True)
    uploaded=st.file_uploader("Attachments (optional)",accept_multiple_files=True,key="u_files")
    files=", ".join([f.name for f in uploaded]) if uploaded else ""
    final=st.text_area("Anything else I should know?",placeholder="Context, history, sensitivities, dependencies.",key="u_final")
    # ── Optional checklist (supervisor can outline steps they want tracked) ──
    st.markdown("**Optional: add a checklist of steps or deliverables**")
    st.caption("If this request involves multiple steps, list them here. They'll appear as a trackable checklist.")
    use_checklist=st.checkbox("Add a checklist to this request",value=False,key="u_use_cl")
    checklist_items=[]
    if use_checklist:
        if "cl_count" not in st.session_state: st.session_state.cl_count=3
        for i in range(st.session_state.cl_count):
            ci=st.text_input(f"Step {i+1}",key=f"cl_item_{i}",placeholder=f"e.g., Book venue, Send invitations, Order catering")
            if ci.strip(): checklist_items.append(ci.strip())
        if st.button("➕ Add another step",key="cl_more"):
            st.session_state.cl_count+=1; st.rerun()
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
        if checklist_items:
            data["Checklist"]=json.dumps([{"text":item,"done":False} for item in checklist_items])
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
def generate_ai_checklist(sub):
    """Use GenAI (Claude) to generate a practical checklist of subtasks for this request.
    Falls back gracefully if no API key or if the call fails."""
    api_key=_s("ANTHROPIC_API_KEY","")
    if not api_key: return None
    try:
        import requests as _req
        rt=sub.get("Request Type","")
        name=sub.get("Your Name","")
        details=sub.get("Additional Notes","")
        deadline=sub.get("Deadline","")
        # Pull in any branch-specific details.
        dj=sub.get("Details JSON","")
        extra=""
        if dj:
            try:
                ex=json.loads(dj)
                extra="\n".join([f"{k}: {v}" for k,v in ex.items() if v and str(v) not in ("","—","False")])
            except: pass
        # Check if supervisor provided their own checklist.
        sup_cl=sub.get("Checklist","")
        sup_note=""
        if sup_cl:
            try:
                items=json.loads(sup_cl)
                sup_note="The supervisor provided these steps: "+", ".join([i["text"] for i in items])+". Incorporate and expand on these."
            except: pass
        prompt=(f"You are an executive administrative assistant at Berea College. "
                f"A supervisor ({name}) submitted a {rt} request with deadline {deadline}. "
                f"Details: {details}\n{extra}\n{sup_note}\n\n"
                f"Generate a practical checklist of 5-10 specific action items to complete this request. "
                f"Each item should be a concrete, actionable step. "
                f"Return ONLY a JSON array of strings, no other text. Example: [\"Book the room\",\"Send invitations\"]")
        resp=_req.post("https://api.anthropic.com/v1/messages",
            headers={"Content-Type":"application/json","x-api-key":api_key,"anthropic-version":"2023-06-01"},
            json={"model":"claude-sonnet-4-6","max_tokens":500,"messages":[{"role":"user","content":prompt}]},
            timeout=15)
        if resp.status_code==200:
            data=resp.json()
            text=data.get("content",[{}])[0].get("text","")
            # Parse the JSON array from the response.
            text=text.strip()
            if text.startswith("```"): text=text.split("```")[1].strip()
            if text.startswith("json"): text=text[4:].strip()
            items=json.loads(text)
            if isinstance(items,list) and all(isinstance(i,str) for i in items):
                return [{"text":item,"done":False} for item in items]
    except Exception:
        pass
    return None

def _extract_event_time(sub):
    """Try to extract a start time from the submission's details. Looks for fields
    like 'Proposed Time', 'Start Time', 'Time', etc. Returns an hour string like '08:00'
    or the default '08:00' if nothing is found."""
    import re
    # Check branch-specific fields from Details JSON.
    dj=sub.get("Details JSON","")
    time_fields=["Proposed Time","Start Time","Time","Start and end time","Time Needed"]
    if dj:
        try:
            ex=json.loads(dj)
            for tf in time_fields:
                val=str(ex.get(tf,"")).strip()
                if val:
                    # Try to extract a time like "3:00 PM", "15:00", "8:45am", etc.
                    m=re.search(r'(\d{1,2}):?(\d{2})?\s*(am|pm|AM|PM)?',val)
                    if m:
                        h=int(m.group(1)); mins=m.group(2) or "00"; ampm=(m.group(3) or "").lower()
                        if ampm=="pm" and h<12: h+=12
                        if ampm=="am" and h==12: h=0
                        return f"{h:02d}:{mins}"
        except: pass
    return "08:00"  # Default

def accept_request(sub):
    """Accept a request: mark Accepted, create a linked task, notify submitter."""
    ref=sub.get("ref_number","")
    name=sub.get("Your Name","")
    rt=sub.get("Request Type","").split("  ")[-1] if "  " in sub.get("Request Type","") else sub.get("Request Type","")
    urg=sub.get("Urgency","")
    old_status=sub.get("Status","New")
    DB.update_status(ref,"Accepted")
    DB.log_audit(ref,"accepted",old_status,"Accepted",cur_user(),st.session_state.get("user_role","admin"))
    # Generate AI checklist (or use supervisor's manual one as fallback).
    checklist=generate_ai_checklist(sub)
    if not checklist:
        sup_cl=sub.get("Checklist","")
        if sup_cl:
            try: checklist=json.loads(sup_cl)
            except: checklist=None
    cl_json=json.dumps(checklist) if checklist else ""
    # Extract calendar-ready datetime for Power Automate.
    deadline=sub.get("Deadline","")
    event_time=_extract_event_time(sub)
    cal_start=f"{deadline}T{event_time}:00" if deadline else ""
    # Default event is 1 hour long.
    try:
        from datetime import datetime as _dt, timedelta as _td
        start_dt=_dt.strptime(cal_start,"%Y-%m-%dT%H:%M:%S")
        end_dt=start_dt+_td(hours=1)
        cal_end=end_dt.strftime("%Y-%m-%dT%H:%M:%S")
    except:
        cal_end=f"{deadline}T{event_time.replace(':','')[:2]}:59:00" if deadline else ""
    DB.save_task({"task_id":gen_task_id(),
        "created_at":datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "Title":f"{rt} — {name}","Category":"Follow-up / admin",
        "Priority":("🔴 Urgent" if "Urgent" in urg else "🟠 High" if "High" in urg else "🟡 Medium"),
        "Deadline":deadline,"Status":"In Progress","Recurrence":"None — one-time",
        "Details":f"Accepted from request {ref} ({name}, {sub.get('Your Role','')}).\n\n{sub.get('Additional Notes','')}",
        "source_ref":ref,"Visibility":"Shared","Checklist":cl_json,
        "Calendar_Start":cal_start,"Calendar_End":cal_end})
    se=sub.get("Notification Email","") or sub.get("Contact Info","")
    notify_accepted(ref,rt,name,se)

def deny_request(sub,reason=""):
    """Deny a request: mark Declined, notify submitter (no task created)."""
    ref=sub.get("ref_number","")
    name=sub.get("Your Name","")
    rt=sub.get("Request Type","").split("  ")[-1] if "  " in sub.get("Request Type","") else sub.get("Request Type","")
    old_status=sub.get("Status","New")
    DB.update_status(ref,"Declined")
    DB.log_audit(ref,"declined",old_status,"Declined",cur_user(),st.session_state.get("user_role","admin"))
    se=sub.get("Notification Email","") or sub.get("Contact Info","")
    notify_declined(ref,rt,name,se,reason)

# ── Shared: persistent nav cards (top of every admin page) ───────────────────
def admin_nav_cards(current=""):
    """Render the five nav cards + navigation buttons at the top of any admin page."""
    subs=DB.get_submissions()
    tasks=DB.get_tasks()
    converted_refs={t.get("source_ref","") for t in tasks if t.get("source_ref","").strip()}
    active_tasks=[t for t in tasks if t.get("Status","") not in ("Complete","")]
    pending_reqs=[s for s in subs
                  if s.get("Status","") in ("New","Under Review","")
                  and s.get("ref_number","") not in converted_refs]
    urg_ct=(sum(1 for s in pending_reqs if "Urgent" in s.get("Urgency","") or "High" in s.get("Urgency",""))
            +sum(1 for t in active_tasks if "Urgent" in t.get("Priority","") or "High" in t.get("Priority","")))
    done_ct=sum(1 for t in tasks if t.get("Status","")=="Complete")
    def _card(col,num,label,cs,btn_label,target_view):
        with col:
            st.markdown(f'<div class="mc"><div class="mn {cs}">{num}</div><div class="ml">{label}</div></div>',unsafe_allow_html=True)
            here=(target_view==current)
            if st.button(("\u2022 You are here" if here else btn_label),key=f"nav_{target_view}_{current}",use_container_width=True,disabled=here):
                st.session_state.view=target_view; st.rerun()
    c0,c1,c2,c3,c4=st.columns(5)
    with c0:
        st.markdown(f'<div class="mc"><div class="mn bl">\U0001f3e0</div><div class="ml">Dashboard</div></div>',unsafe_allow_html=True)
        is_home=(current in ("admin","dashboard"))
        if st.button(("\u2022 You are here" if is_home else "Dashboard \u2192"),key=f"nav_dashboard_{current}",use_container_width=True,disabled=is_home):
            st.session_state.view="admin"; st.rerun()
    _card(c1,len(active_tasks),"On my plate","bl","My tasks \u2192","tasks")
    _card(c2,len(pending_reqs),"Pending acceptance","bl","Review \u2192","pending")
    _card(c3,urg_ct,"Urgent / high","am","View urgent \u2192","urgent")
    _card(c4,done_ct,"Completed","gr","Completed \u2192","completed")

# ── View: Admin dashboard ────────────────────────────────────────────────────
def v_admin():
    banner("Admin Dashboard — Everything on My Plate","Berea College · Office of the Provost · Admin view")
    subs=DB.get_submissions()
    tasks=DB.get_tasks()
    admin_nav_cards(current="admin")
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
                _render_request_details(sub)
            with cr:
                # Has this request already been converted to a task?
                existing_tasks=DB.get_tasks()
                already_converted=any(t.get("source_ref","")==ref for t in existing_tasks)

                if not already_converted and stat not in ["Complete","Declined","Cancelled"]:
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
    admin_nav_cards(current="pending")
    st.markdown("<br>",unsafe_allow_html=True)
    subs=DB.get_submissions()
    tasks=DB.get_tasks()
    converted_refs={t.get("source_ref","") for t in tasks if t.get("source_ref","").strip()}
    pending=[s for s in subs if s.get("Status","") in ("New","Under Review","")
             and s.get("ref_number","") not in converted_refs]
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
            _render_request_details(sub)
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
    admin_nav_cards(current="urgent")
    st.markdown("<br>",unsafe_allow_html=True)
    subs=DB.get_submissions()
    tasks=DB.get_tasks()
    converted_refs={t.get("source_ref","") for t in tasks if t.get("source_ref","").strip()}
    u_tasks=[t for t in tasks if t.get("Status","") not in ("Complete","")
             and ("Urgent" in t.get("Priority","") or "High" in t.get("Priority",""))]
    u_reqs=[s for s in subs if s.get("Status","") in ("New","Under Review","")
            and s.get("ref_number","") not in converted_refs
            and ("Urgent" in s.get("Urgency","") or "High" in s.get("Urgency",""))]
    st.caption(f"{len(u_tasks)} urgent/high task{'s' if len(u_tasks)!=1 else ''} and "
               f"{len(u_reqs)} urgent/high request{'s' if len(u_reqs)!=1 else ''} awaiting acceptance.")
    if not u_tasks and not u_reqs:
        st.success("Nothing urgent right now."); return
    if u_tasks:
        sec("🔴 Urgent / high tasks")
        for t in sorted(u_tasks,key=lambda x:x.get("Deadline","")):
            tid,pri,uc,stat,vis,rec=_task_props(t)
            with st.expander(f"{pri.split(' ')[0]} {t.get('Title','')}  ·  {t.get('Category','')}  ·  due {t.get('Deadline','')}"):
                _render_task_detail(t,tid,stat,pri,uc,rec,vis)
    if u_reqs:
        sec("🟠 Urgent / high requests (awaiting acceptance)")
        for s in sorted(u_reqs,key=lambda x:x.get("Deadline","")):
            ref=s.get("ref_number",""); name=s.get("Your Name","")
            urg=s.get("Urgency",""); uc_r=ucss(urg)
            rt=s.get("Request Type","").split("  ")[-1] if "  " in s.get("Request Type","") else s.get("Request Type","")
            with st.expander(f"{'🔴 ' if 'Urgent' in urg else '🟠 '}{name}  ·  {rt}  ·  due {s.get('Deadline','')}"):
                _render_request_details(s)
                b1,b2=st.columns(2)
                with b1:
                    if st.button("✅ Accept",key=f"uacc_{ref}",use_container_width=True):
                        accept_request(s); st.success("Accepted."); st.rerun()
                with b2:
                    if st.button("🚫 Deny",key=f"uden_{ref}",use_container_width=True):
                        deny_request(s); st.warning("Declined."); st.rerun()

# ── View: Supervisor portal ───────────────────────────────────────────────────
def v_portal():
    user=cur_user()
    # Try to get role from users sheet first, fall back to secrets.
    u_record=DB.get_user_by_email(st.session_state.get("user_email",""))
    role_t=u_record.get("title","") if u_record else SUPERVISORS.get(user,{}).get("role","")
    banner(f"Support Requests",f"Berea College · {role_t} · {user}")
    # ── View toggle: My requests vs All supervisors ──────────────────────
    vt1,vt2=st.columns([2,3])
    with vt1:
        view_scope=st.selectbox("Viewing",["My requests only","All supervisors' requests"],key="sup_scope")
    show_all=(view_scope!="My requests only")
    if show_all:
        subs=DB.get_submissions()
        st.caption(f"Showing requests from **all supervisors**. Toggle above to see only yours.")
    else:
        subs=DB.get_submissions(name=user)
        st.caption(f"Showing only **your** requests. Toggle above to see all supervisors.")
    total=len(subs)
    # Each request falls into exactly ONE bucket below (no overlap).
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
    cancelled=sum(1 for s in subs if s.get("Status","") in ("Cancelled","Declined"))
    scope_label="across all supervisors" if show_all else "from you"
    st.caption(f"**{total}** request{'s' if total!=1 else ''} {scope_label}."
               f"{f' ({cancelled} cancelled/declined — still in the list for records.)' if cancelled else ''}")
    st.markdown(f'<div class="ro-notice" style="margin-top:.6rem;">👁️ Requests are <strong>read-only</strong> once submitted. To flag a change or update, add a note.</div>',unsafe_allow_html=True)
    pii_banner(compact=True)
    if not subs: st.info("No requests to show for this view." if show_all else "No requests yet. Use the sidebar to submit your first request."); return
    # ── Timeline ────────────────────────────────────────────────────────────
    st.markdown("<br>",unsafe_allow_html=True)
    sec("📊 Timeline (Gantt view)")
    admin_opts=["All",ADMIN_NAME]
    hcol1,hcol2=st.columns([2,3])
    with hcol1:
        handled_by=st.selectbox("Handled by",admin_opts,key="sup_gantt_handler")
    if show_all:
        # Show all supervisors' requests, colored by supervisor name.
        gantt_data=[]
        for s in subs:
            start=_parse_date(s.get("submitted_at","")); end=_parse_date(s.get("Deadline",""))
            if not start or not end: continue
            if end<start: end=start
            gantt_data.append({"title":s.get("Request Type","").split("  ")[-1] if "  " in s.get("Request Type","") else s.get("Request Type",""),
                              "start":start,"end":end,"status":s.get("Status","New"),
                              "group":s.get("Your Name","Unknown"),"complete":s.get("Status","") in ("Complete","Cancelled","Declined")})
        draw_gantt(gantt_data,legend_label="SUPERVISOR")
    else:
        st.caption("Your requests drawn to their deadlines. Colored by request type; completed items are greyed out.")
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
        is_done=stat in ("Complete","Cancelled","Declined")
        nct=len(DB.get_notes(ref)); short=ref.split("-")[-1] if ref else "—"
        fstat=friendly_status(stat)
        who_prefix=f'{sub.get("Your Name","")} · ' if show_all else ""
        done_marker="~~" if is_done else ""
        with st.expander(f"{done_marker}{who_prefix}{short}  ·  {rt}  ·  {fstat}  ·  {f'{nct} note(s)' if nct else 'No notes'}{done_marker}"):
            _render_request_details(sub)
            # Cancel / recall button (supervisor can cancel their own request only).
            is_own=(sub.get("Your Name","")==user)
            if is_own and stat not in ("Complete","Cancelled","Declined"):
                if st.button(f"❌ Cancel this request",key=f"cancel_{ref}"):
                    DB.update_status(ref,"Cancelled")
                    DB.log_audit(ref,"cancelled",stat,"Cancelled",cur_user(),"supervisor")
                    st.warning("Request cancelled. It remains in your records for reference.")
                    st.rerun()
            # Audit trail (collapsible).
            trail=DB.get_audit(ref)
            if trail:
                with st.expander(f"📜 History ({len(trail)} event{'s' if len(trail)!=1 else ''})"):
                    for entry in trail:
                        st.markdown(f'<div style="font-size:.75rem;color:#5a7a96;border-bottom:1px solid #eee;padding:4px 0;">'
                                    f'{entry.get("timestamp","")} — <strong>{entry.get("action","")}</strong>'
                                    f'{" · " + entry.get("old_value","") + " → " + entry.get("new_value","") if entry.get("new_value") else ""}'
                                    f' — by {entry.get("by_user","")}</div>',unsafe_allow_html=True)
            hr(); notes_thread(ref,complete=(stat in ("Complete","Cancelled")))

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
                # Read-only progress tracker for supervisors.
                cl_raw=tk.get("Checklist","")
                cl_items=None
                if cl_raw:
                    try: cl_items=json.loads(cl_raw)
                    except: cl_items=None
                if cl_items and isinstance(cl_items,list):
                    total_cl=len(cl_items); done_cl=sum(1 for c in cl_items if c.get("done",False))
                    pct=int(done_cl/total_cl*100) if total_cl else 0
                    pct_color="#27500A" if pct==100 else "#004175" if pct>50 else "#C75B12"
                    st.markdown(f'<div style="margin:6px 0;font-size:.8rem;font-weight:700;color:{pct_color};">'
                                f'📋 Progress: {done_cl}/{total_cl} steps ({pct}%)</div>',unsafe_allow_html=True)
                    st.progress(pct/100)
                    for item in cl_items:
                        icon="✅" if item.get("done",False) else "⬜"
                        st.markdown(f'{icon} {item.get("text","")}')
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
        pii_banner(compact=True)
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
    Bars grey out when complete. Colored by 'group'. Dates shown on each bar."""
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
    span=max((max_d-min_d).days,1); pad=max(int(span*0.06),3)
    min_d=min_d-_td2(days=pad); max_d=max_d+_td2(days=pad)
    total_days=max((max_d-min_d).days,1)
    rows=sorted(rows,key=lambda r:(r["start"],r["end"]))
    row_h=36; gap=10; top=58; left=280; right=40
    width=1020; plot_w=width-left-right
    # Extra space for legend row below.
    legend_h=32
    height=top+len(rows)*(row_h+gap)+legend_h+10
    def x_of(d): return left+plot_w*((d-min_d).days/total_days)
    parts=[f'<svg viewBox="0 0 {width} {height}" xmlns="http://www.w3.org/2000/svg" '
           f'style="width:100%;max-width:1020px;height:auto;font-family:Barlow,Calibri,sans-serif;">']
    parts.append(f'<rect x="0" y="0" width="{width}" height="{height}" fill="#ffffff"/>')
    # Month gridlines + labels.
    cur=_dt(min_d.year,min_d.month,1)
    while cur<max_d:
        gx=x_of(cur)
        if gx>=left:
            parts.append(f'<line x1="{gx:.1f}" y1="{top-10}" x2="{gx:.1f}" y2="{height-legend_h-5}" stroke="#e3e9f0" stroke-width="1"/>')
            parts.append(f'<text x="{gx+4:.1f}" y="{top-16}" font-size="11" fill="#5a7a96" font-weight="600">{cur.strftime("%b %Y")}</text>')
        cur=_dt(cur.year+1,1,1) if cur.month==12 else _dt(cur.year,cur.month+1,1)
    # "Today" line.
    today=_dt.now()
    if min_d<=today<=max_d:
        tx=x_of(today)
        parts.append(f'<line x1="{tx:.1f}" y1="{top-10}" x2="{tx:.1f}" y2="{height-legend_h-5}" stroke="#C75B12" stroke-width="2" stroke-dasharray="4 3"/>')
        parts.append(f'<text x="{tx+4:.1f}" y="{height-legend_h-8}" font-size="9.5" fill="#C75B12" font-weight="700">Today</text>')
    # Bars.
    for i,r in enumerate(rows):
        y=top+i*(row_h+gap)
        bx=x_of(r["start"]); bw=max(x_of(r["end"])-bx,8)
        fill="#c9ced6" if r["complete"] else color_of[r["group"]]
        opacity="0.45" if r["complete"] else "1"
        txt_color="#888" if r["complete"] else "#1a1a1a"
        # Task label (left, truncated to avoid overlap with bar area).
        max_chars=32
        label=r["title"] if len(r["title"])<=max_chars else r["title"][:max_chars-1]+"…"
        deco=' text-decoration="line-through"' if r["complete"] else ''
        parts.append(f'<text x="8" y="{y+row_h*0.55:.1f}" font-size="11.5" fill="{txt_color}"{deco}>{label}</text>')
        # Date range below the label (small, non-overlapping).
        date_str=f'{r["start"].strftime("%b %d")} → {r["end"].strftime("%b %d")}'
        parts.append(f'<text x="8" y="{y+row_h*0.88:.1f}" font-size="8.5" fill="#8a9ab0">{date_str}</text>')
        # Bar.
        parts.append(f'<rect x="{bx:.1f}" y="{y+2:.1f}" width="{bw:.1f}" height="{row_h-4}" rx="5" fill="{fill}" opacity="{opacity}"/>')
        # Status tag on the bar if wide enough.
        tag=friendly_status(r["status"]) if not r["complete"] else "Complete"
        short_tag={"Submitted & pending confirmation of availability":"Pending","Confirmed & underway":"Accepted",
                   "In progress":"In progress","Awaiting your input":"Awaiting","Under review by Tiff":"Review",
                   "On hold":"On hold","Complete":"Done","Cancelled by you":"Cancelled",
                   "Not able to take this on":"Declined"}.get(tag,tag[:12])
        if bw>60:
            parts.append(f'<text x="{bx+7:.1f}" y="{y+row_h*0.55:.1f}" font-size="10" fill="#ffffff" font-weight="600">{short_tag}</text>')
            # End-date on the bar.
            end_lbl=r["end"].strftime("%b %d")
            if bw>120:
                parts.append(f'<text x="{bx+bw-5:.1f}" y="{y+row_h*0.55:.1f}" font-size="9" fill="rgba(255,255,255,0.8)" text-anchor="end">{end_lbl}</text>')
    # Legend row.
    ly=height-10; lx=left
    parts.append(f'<text x="8" y="{ly:.1f}" font-size="10" fill="#5a7a96" font-weight="700">{legend_label}:</text>')
    for g in groups:
        c=color_of[g]
        parts.append(f'<rect x="{lx:.1f}" y="{ly-9:.1f}" width="11" height="11" rx="2" fill="{c}"/>')
        gl=g if len(g)<=20 else g[:18]+"…"
        parts.append(f'<text x="{lx+15:.1f}" y="{ly:.1f}" font-size="10" fill="#333">{gl}</text>')
        lx+=22+len(gl)*6
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

def add_task_form():
    st.markdown("<br>",unsafe_allow_html=True)
    sec("➕ Add a New Task")
    pii_banner(compact=True)
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
                            help="On by default. Uncheck to keep this off supervisors' dashboards.")
        if st.form_submit_button("➕ Add task"):
            if not t_title.strip():
                st.error("⚠️ Task title is required.")
            else:
                tid=gen_task_id()
                DB.save_task({"task_id":tid,"created_at":datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "Title":t_title.strip(),"Category":t_cat,"Priority":t_pri,
                    "Deadline":t_ddl.strftime("%Y-%m-%d"),"Status":"New","Recurrence":t_rec,
                    "Details":t_det.strip(),"source_ref":"","Visibility":"Shared" if t_share else "Private",
                    "Owner":cur_user(),"AssignedTo":cur_user(),"StaffVisible":"Yes"})
                DB.log_audit(tid,"task_created","",t_title.strip(),cur_user(),st.session_state.get("user_role",""))
                st.success(f"Task added: {t_title.strip()}"); st.rerun()

def _render_request_details(sub):
    """Render a request's details in a clean, readable format. Parses the Details
    JSON into labeled fields instead of dumping raw JSON."""
    ref=sub.get("ref_number",""); name=sub.get("Your Name","")
    urg=sub.get("Urgency",""); uc=ucss(urg)
    stat=sub.get("Status","New"); sc=S_CSS.get(stat,"s-new")
    fstat=friendly_status(stat)
    conf="Yes" in str(sub.get("Confidential",""))
    cf='<span style="color:#c00;font-weight:700;">🔒 CONFIDENTIAL</span>' if conf else ""
    st.markdown(f'<div style="font-size:.82rem;">'
                f'<span class="ref-code">{ref}</span> {cf}<br>'
                f'<strong>Submitted by:</strong> {name} ({sub.get("Your Role","")})<br>'
                f'<strong>Contact:</strong> {sub.get("Contact Info","")}<br>'
                f'<strong>Type:</strong> {sub.get("Request Type","")}<br>'
                f'<strong>Urgency:</strong> <span class="sb {uc}">{urg}</span> &nbsp; '
                f'<strong>Deadline:</strong> {sub.get("Deadline","")}<br>'
                f'<strong>Status:</strong> <span class="sb {sc}">{fstat}</span><br>'
                f'<strong>Submitted:</strong> {sub.get("submitted_at","")}'
                f'</div>',unsafe_allow_html=True)
    # Parse the Details JSON into readable fields (the important fix).
    dj=sub.get("Details JSON","")
    if dj:
        try:
            ex=json.loads(dj)
            if ex:
                # Filter out empty/boring fields and internal keys.
                skip_keys={"ref_number","submitted_at","Status","Details JSON","Your Name","Your Role",
                           "Contact Info","Deadline","Urgency","Confidential","Request Type",
                           "Notify on Status Change","Notify on New Note","Notification Email",
                           "File Attachments","Checklist"}
                items=[(k,v) for k,v in ex.items() if v and str(v) not in ("","—","False","{}","[]","0") and k not in skip_keys]
                if items:
                    st.markdown("<strong style='font-size:.82rem;'>Request details:</strong>",unsafe_allow_html=True)
                    for k,v in items:
                        st.markdown(f'<div style="font-size:.8rem;padding:2px 0;"><strong>{k}:</strong> {v}</div>',unsafe_allow_html=True)
        except: pass
    # Show additional notes if present.
    notes=sub.get("Additional Notes","")
    if notes:
        st.markdown(f'<div style="font-size:.82rem;margin-top:8px;background:#f5f8fc;border-radius:6px;padding:8px 11px;">'
                    f'<strong>Additional notes:</strong><br>{notes}</div>',unsafe_allow_html=True)
    # Show file attachments if present.
    files=sub.get("File Attachments","")
    if files:
        st.markdown(f'<div style="font-size:.8rem;margin-top:4px;"><strong>Attachments:</strong> {files}</div>',unsafe_allow_html=True)

def _render_task_detail(t,tid,stat,pri,uc,rec,vis):
    """Render the detail panel for a single task (shared by list and card views)."""
    rec_badge=f' &nbsp; <span class="sb s-review">\U0001f501 {rec}</span>' if rec not in ["","None \u2014 one-time"] else ""
    src=t.get("source_ref","")
    src_badge=f' &nbsp; <span class="sb s-hold">from {src.split("-")[-1]}</span>' if src else ""
    vis_badge=(' &nbsp; <span class="sb s-done">\U0001f441\ufe0f Shared</span>' if vis!="Private"
               else ' &nbsp; <span class="sb s-hold">\U0001f512 Private</span>')
    st.markdown(f'<div style="font-size:.82rem;"><strong>Category:</strong> {t.get("Category","")}<br>'
                f'<strong>Priority:</strong> <span class="sb {uc}">{pri}</span> &nbsp; '
                f'<strong>Deadline:</strong> {t.get("Deadline","")}{rec_badge}{src_badge}{vis_badge}<br>'
                f'<strong>Created:</strong> {t.get("created_at","")}</div>',unsafe_allow_html=True)
    if t.get("Details",""): st.markdown(f'<div style="font-size:.84rem;margin-top:6px;background:#f5f8fc;border-radius:6px;padding:8px 11px;">{t.get("Details","")}</div>',unsafe_allow_html=True)
    cl_raw=t.get("Checklist",""); cl_items=None
    if cl_raw:
        try: cl_items=json.loads(cl_raw)
        except: cl_items=None
    if cl_items and isinstance(cl_items,list):
        total_cl=len(cl_items); done_cl=sum(1 for c in cl_items if c.get("done",False))
        pct=int(done_cl/total_cl*100) if total_cl else 0
        pct_color="#27500A" if pct==100 else "#004175" if pct>50 else "#C75B12"
        st.markdown(f'<div style="margin:10px 0 6px;font-size:.8rem;font-weight:700;color:{pct_color};">\U0001f4cb Checklist: {done_cl}/{total_cl} ({pct}%)</div>',unsafe_allow_html=True)
        st.progress(pct/100)
    editable=can_edit_task(t)
    assignee=str(t.get("AssignedTo","")).strip()
    if cl_items and isinstance(cl_items,list):
        if editable:
            cl_changed=False
            for ci,item in enumerate(cl_items):
                new_val=st.checkbox(item.get("text",""),value=item.get("done",False),key=f"cl_{tid}_{ci}")
                if new_val!=item.get("done",False):
                    cl_items[ci]["done"]=new_val; cl_changed=True
            if cl_changed:
                DB.update_task(tid,"Checklist",json.dumps(cl_items))
                DB.log_audit(tid,"checklist_updated","",f"{sum(1 for c in cl_items if c.get('done'))}/{len(cl_items)} done",cur_user(),st.session_state.get("user_role",""))
                st.rerun()
        else:
            # Read-only checklist for non-owners.
            for item in cl_items:
                icon="✅" if item.get("done",False) else "⬜"
                st.markdown(f'{icon} {item.get("text","")}')
    hr()
    if not editable:
        # Notes-only view for people who don't own/aren't assigned this task.
        who=assignee or t.get("Owner","") or "another staff member"
        st.markdown(f'<div class="ro-notice">🔒 This task is assigned to <strong>{who}</strong>. '
                    f'You can view it and add notes, but only the assigned person can make changes.</div>',unsafe_allow_html=True)
    else:
        uc1,uc2,uc3=st.columns(3)
        with uc1: ns=st.selectbox("Status",STATUS_OPTIONS,index=STATUS_OPTIONS.index(stat) if stat in STATUS_OPTIONS else 0,key=f"ts_{tid}")
        with uc2: np=st.selectbox("Priority",TASK_PRIORITIES,index=TASK_PRIORITIES.index(pri) if pri in TASK_PRIORITIES else 2,key=f"tp_{tid}")
        with uc3: nv=st.selectbox("Visibility",["Shared","Private"],index=0 if vis!="Private" else 1,key=f"tv_{tid}")
        if st.button("\U0001f4be Update",key=f"tu_{tid}"):
            if ns!=stat:
                DB.update_task(tid,"Status",ns)
                DB.log_audit(tid,"status_change",stat,ns,cur_user(),st.session_state.get("user_role",""))
                rec_val=t.get("Recurrence","")
                if ns=="Complete" and rec_val not in ["","None \u2014 one-time"]:
                    nd=next_occurrence(t.get("Deadline",""),rec_val)
                    DB.save_task({"task_id":gen_task_id(),"created_at":datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "Title":t.get("Title",""),"Category":t.get("Category",""),
                        "Priority":pri,"Deadline":nd,"Status":"New","Recurrence":rec_val,
                        "Details":t.get("Details",""),"source_ref":"","Visibility":t.get("Visibility","Shared"),
                        "Owner":t.get("Owner",""),"AssignedTo":t.get("AssignedTo",""),"StaffVisible":t.get("StaffVisible","")})
                    st.success(f"Completed. Next occurrence created for {nd}.")
            if np!=pri:
                DB.update_task(tid,"Priority",np)
                DB.log_audit(tid,"priority_change",pri,np,cur_user(),st.session_state.get("user_role",""))
            if nv!=vis:
                DB.update_task(tid,"Visibility",nv)
                DB.log_audit(tid,"visibility_change",vis,nv,cur_user(),st.session_state.get("user_role",""))
            st.rerun()
        with st.expander("\U0001f5d1\ufe0f Delete this task"):
            st.caption("This permanently removes the task.")
            if st.button("Confirm delete",key=f"td_{tid}"):
                DB.log_audit(tid,"task_deleted",t.get("Title",""),"",cur_user(),st.session_state.get("user_role",""))
                DB.delete_task(tid); st.rerun()
    # Admin-only hidden activity log.
    if is_admin():
        log=DB.get_audit(tid)
        with st.expander(f"🔍 Activity log ({len(log)} event{'s' if len(log)!=1 else ''}) — admin only"):
            if not log:
                st.caption("No recorded activity for this task yet.")
            else:
                for entry in sorted(log,key=lambda x:x.get("timestamp",""),reverse=True):
                    chg=f' · {entry.get("old_value","")} → {entry.get("new_value","")}' if entry.get("new_value") else ""
                    st.markdown(f'<div style="font-size:.74rem;color:#5a7a96;border-bottom:1px solid #eee;padding:3px 0;">'
                                f'{entry.get("timestamp","")} — <strong>{entry.get("action","")}</strong>{chg} '
                                f'— by {entry.get("by_user","")} ({entry.get("by_role","")})</div>',unsafe_allow_html=True)
    # Quick actions row.
    qa1,qa2,qa3=st.columns(3)
    with qa1:
        if st.button("\U0001f50d Full detail view",key=f"qd_{tid}",use_container_width=True):
            st.session_state.detail_task_id=tid; st.session_state.view="task_detail"; st.rerun()
    with qa2:
        ics=generate_ics(t)
        if ics:
            st.download_button("\U0001f4c5 Add to calendar",data=ics,file_name=f"task_{tid}.ics",
                              mime="text/calendar",key=f"ics_{tid}",use_container_width=True)
    with qa3:
        if st.button("\U0001f4dd Summary",key=f"qs_{tid}",use_container_width=True):
            summary=generate_task_summary(t)
            st.session_state[f"summary_{tid}"]=summary; st.rerun()
    # Show cached summary if it exists.
    cached_summary=st.session_state.get(f"summary_{tid}","")
    if cached_summary:
        st.markdown(f'<div style="background:#f5f8fc;border-radius:6px;padding:10px 14px;font-size:.84rem;margin:6px 0;">{cached_summary}</div>',unsafe_allow_html=True)
    hr(); task_notes_thread(tid,complete=(stat=="Complete"))

def _task_props(t):
    tid=t.get("task_id",""); pri=t.get("Priority","")
    uc=("urg-u" if "Urgent" in pri else "urg-h" if "High" in pri else "urg-s" if "Medium" in pri else "urg-l")
    stat=t.get("Status","New"); vis=str(t.get("Visibility","Shared")); rec=t.get("Recurrence","")
    return tid,pri,uc,stat,vis,rec

def visible_tasks(all_tasks,viewer_name):
    """Everyone (admin, colleague, support, supervisor) can see all tasks —
    full transparency. Filtering by owner is offered in the UI, not enforced here."""
    return all_tasks

def can_edit_task(t):
    """Determine whether the current user may directly edit this task.
    - Admins can always edit.
    - Legacy tasks with no recorded owner stay editable by anyone (nothing to lock to).
    - Before a task is accepted (New/Under Review): the owner or assignee can edit.
    - Once accepted or later: only the assigned person (or an admin) can edit."""
    if is_admin(): return True
    me=cur_user()
    owner=str(t.get("Owner","")).strip()
    assignee=str(t.get("AssignedTo","")).strip()
    if not owner and not assignee: return True  # legacy task, nothing to lock to
    status=t.get("Status","New")
    accepted=status not in ("New","Under Review","")
    if accepted:
        if assignee: return assignee==me
        return owner==me
    return me in (owner,assignee)

def v_tasks():
    # Heading adapts to role.
    if is_admin():
        banner("My Tasks","Berea College · Office of the Provost · All tasks")
    elif is_colleague():
        banner("Tasks","Berea College · Office of the Provost · Your tasks + shared team tasks")
    else:
        banner("My Tasks","Berea College · Office of the Provost · Your assigned tasks")
    admin_nav_cards(current="tasks")
    all_tasks_raw=DB.get_tasks()
    # Apply staff-to-staff visibility (admins see all; others see own + shared).
    tasks=visible_tasks(all_tasks_raw,cur_user())
    active=[t for t in tasks if t.get("Status","")!="Complete"]
    done=[t for t in tasks if t.get("Status","")=="Complete"]
    if is_admin():
        st.caption(f"📊 Task diagnostic: {len(all_tasks_raw)} total in storage, {len(active)} active, {len(done)} complete.")
    recurring=sum(1 for t in tasks if t.get("Recurrence","") not in ["","None \u2014 one-time"])
    if recurring: st.caption(f"Of your active tasks, {recurring} recurring auto-repeat when completed.")
    if not tasks:
        st.markdown("<br>",unsafe_allow_html=True)
        st.info("No tasks yet. Use the Add a New Task form at the bottom of this page.")
        add_task_form(); return
    st.markdown("<br>",unsafe_allow_html=True)
    tc1,tc2,tc3,tc4=st.columns([1.2,1.4,1.4,1.2])
    with tc1: view_mode=st.selectbox("View as",["\U0001f4cb List","\U0001f0cf Cards (by status)"],key="task_view_mode",label_visibility="collapsed")
    with tc2: fcat=st.selectbox("Filter by category",["All categories"]+TASK_CATEGORIES,key="tf_cat",label_visibility="collapsed")
    with tc3: fpri=st.selectbox("Filter by priority",["All priorities"]+TASK_PRIORITIES,key="tf_pri",label_visibility="collapsed")
    with tc4: fown=st.selectbox("Whose tasks",["Everyone's tasks","Just mine"],key="tf_own",label_visibility="collapsed")
    shown=active
    if fcat!="All categories": shown=[t for t in shown if t.get("Category","")==fcat]
    if fpri!="All priorities": shown=[t for t in shown if t.get("Priority","")==fpri]
    if fown=="Just mine":
        me=cur_user()
        shown=[t for t in shown if t.get("Owner","")==me or t.get("AssignedTo","")==me or not t.get("Owner","")]
    is_card_view=("Cards" in view_mode)
    if is_card_view:
        status_order=["New","Under Review","Accepted","In Progress","Awaiting Info","On Hold"]
        status_groups={s:[t for t in shown if t.get("Status","")==s] for s in status_order}
        active_statuses=[s for s in status_order if status_groups[s]]
        if not active_statuses: st.caption("No active tasks match your filters.")
        else:
            sec("\U0001f4cb Active Tasks \u2014 Card View")
            cols=st.columns(len(active_statuses))
            for ci,stat_name in enumerate(active_statuses):
                with cols[ci]:
                    sc=S_CSS.get(stat_name,"s-new"); count=len(status_groups[stat_name])
                    st.markdown(f'<div style="background:#f0f5fb;border-radius:8px;padding:8px 10px;margin-bottom:10px;border-top:3px solid {BLUE};text-align:center;"><span class="sb {sc}" style="font-size:.8rem;">{stat_name}</span><div style="font-size:1.4rem;font-weight:700;color:{BLUE};margin:4px 0;">{count}</div></div>',unsafe_allow_html=True)
                    for t in sorted(status_groups[stat_name],key=lambda x:x.get("Deadline","")):
                        tid,pri,uc,stat,vis,rec=_task_props(t)
                        title=t.get("Title","(untitled)"); ddl=t.get("Deadline","")
                        cl_raw=t.get("Checklist",""); pct_html=""
                        if cl_raw:
                            try:
                                cli=json.loads(cl_raw); cl_total=len(cli); dc=sum(1 for c in cli if c.get("done",False))
                                pct=int(dc/cl_total*100) if cl_total else 0
                                pct_html=f'<div style="background:#e0e6ed;border-radius:3px;height:5px;margin:5px 0;"><div style="background:{"#27500A" if pct==100 else "#004175"};width:{pct}%;height:5px;border-radius:3px;"></div></div>'
                            except: pass
                        pri_icon=pri.split(" ")[0] if pri else ""
                        st.markdown(f'<div style="background:#fff;border:1px solid #dde3eb;border-left:4px solid {BLUE};border-radius:0 8px 8px 0;padding:9px 11px;margin-bottom:8px;font-size:.8rem;"><div style="font-weight:700;color:#1a1a1a;">{pri_icon} {title}</div><div style="font-size:.72rem;color:#5a7a96;margin-top:3px;"><span class="sb {uc}">{pri}</span> \u00b7 due {ddl}</div>{pct_html}</div>',unsafe_allow_html=True)
                        with st.expander(f"Edit: {title[:25]}\u2026" if len(title)>25 else f"Edit: {title}",expanded=False):
                            _render_task_detail(t,tid,stat,pri,uc,rec,vis)
    else:
        sec("\U0001f4cb Active Tasks")
        if not shown: st.caption("No active tasks match your filters.")
        for t in sorted(shown,key=lambda x:x.get("Deadline","")):
            tid,pri,uc,stat,vis,rec=_task_props(t)
            with st.expander(f"{pri.split(' ')[0]} {t.get('Title','')}  \u00b7  {t.get('Category','')}  \u00b7  {stat}"):
                _render_task_detail(t,tid,stat,pri,uc,rec,vis)
    st.markdown("<br><br>",unsafe_allow_html=True)
    sec("\u2705 Completed Tasks")
    done_filtered=done
    if fcat!="All categories": done_filtered=[t for t in done_filtered if t.get("Category","")==fcat]
    if not done_filtered:
        st.caption("No completed tasks" + (" matching your filters." if fcat!="All categories" else " yet. They'll appear here as you finish work."))
    else:
        st.caption(f"{len(done_filtered)} completed task{'s' if len(done_filtered)!=1 else ''}. These are greyed out and kept for your records.")
        for t in sorted(done_filtered,key=lambda x:x.get("Deadline",""),reverse=True):
            tid=t.get("task_id",""); title=t.get("Title",""); cat=t.get("Category",""); ddl=t.get("Deadline","")
            src=t.get("source_ref",""); src_note=f" \u00b7 from {src.split('-')[-1]}" if src else ""
            st.markdown(f'<div style="background:#f5f5f5;border:1px solid #e0e0e0;border-left:4px solid #c9ced6;border-radius:0 6px 6px 0;padding:8px 12px;margin-bottom:6px;font-size:.8rem;color:#888;"><span style="text-decoration:line-through;">{title}</span> &nbsp;\u00b7&nbsp; {cat}{src_note} &nbsp;\u00b7&nbsp; {ddl}</div>',unsafe_allow_html=True)
    add_task_form()

# ── Helper: Generate .ics calendar file for a task ───────────────────────────
def generate_ics(task):
    """Generate a downloadable .ics calendar file for a task."""
    title=task.get("Title","Task")
    deadline=task.get("Deadline","")
    details=task.get("Details","").replace("\n","\\n")
    cal_start=task.get("Calendar_Start","")
    cal_end=task.get("Calendar_End","")
    if not cal_start and deadline:
        cal_start=f"{deadline}T08:00:00"
        cal_end=f"{deadline}T09:00:00"
    if not cal_start: return None
    # Format for iCal (remove dashes and colons).
    start_ical=cal_start.replace("-","").replace(":","").replace("T","T")
    end_ical=cal_end.replace("-","").replace(":","").replace("T","T") if cal_end else start_ical
    # 1-week and 1-day reminders.
    ics=f"""BEGIN:VCALENDAR
VERSION:2.0
PRODID:-//Berea College//Support Request System//EN
BEGIN:VEVENT
DTSTART:{start_ical}
DTEND:{end_ical}
SUMMARY:DEADLINE: {title}
DESCRIPTION:{details[:500]}
STATUS:CONFIRMED
BEGIN:VALARM
TRIGGER:-P7D
ACTION:DISPLAY
DESCRIPTION:1 week reminder: {title}
END:VALARM
BEGIN:VALARM
TRIGGER:-P1D
ACTION:DISPLAY
DESCRIPTION:1 day reminder: {title}
END:VALARM
END:VEVENT
END:VCALENDAR"""
    return ics

# ── Helper: AI-generated task summary ────────────────────────────────────────
def generate_task_summary(task):
    """Generate a brief 'done vs. still needed' summary using AI, or manually from checklist."""
    title=task.get("Title","")
    details=task.get("Details","")
    status=task.get("Status","")
    cl_raw=task.get("Checklist","")
    # Try checklist-based summary first (no API needed).
    cl_items=None
    if cl_raw:
        try: cl_items=json.loads(cl_raw)
        except: cl_items=None
    if cl_items and isinstance(cl_items,list):
        done_items=[i["text"] for i in cl_items if i.get("done",False)]
        todo_items=[i["text"] for i in cl_items if not i.get("done",False)]
        total=len(cl_items); done_n=len(done_items)
        pct=int(done_n/total*100) if total else 0
        summary=f"**Progress: {done_n}/{total} ({pct}%)**\n\n"
        if done_items:
            summary+="**Completed:**\n"+"".join(f"- ✅ {i}\n" for i in done_items)
        if todo_items:
            summary+="\n**Still needed:**\n"+"".join(f"- ⬜ {i}\n" for i in todo_items)
        return summary
    # If no checklist, try AI summary.
    api_key=_s("ANTHROPIC_API_KEY","")
    if api_key:
        try:
            import requests as _req
            notes=DB.get_notes(task.get("task_id",""))
            notes_text="\n".join([f"- {n.get('author','')}: {n.get('text','')}" for n in notes[-5:]]) if notes else "No notes yet."
            prompt=(f"You are an executive admin assistant. Give a 3-4 sentence summary of this task's status.\n"
                    f"Task: {title}\nStatus: {status}\nDetails: {details}\nRecent notes:\n{notes_text}\n\n"
                    f"Format: first sentence = overall status, then what's been done, then what still needs attention.")
            resp=_req.post("https://api.anthropic.com/v1/messages",
                headers={"Content-Type":"application/json","x-api-key":api_key,"anthropic-version":"2023-06-01"},
                json={"model":"claude-sonnet-4-6","max_tokens":300,"messages":[{"role":"user","content":prompt}]},
                timeout=15)
            if resp.status_code==200:
                return resp.json().get("content",[{}])[0].get("text","")
        except: pass
    # Fallback: simple status-based summary.
    return f"**{title}** is currently **{status}**. {details[:200] if details else 'No additional details recorded.'}"

# ── View: Today's Focus (consolidated daily checklist) ───────────────────────
def v_today():
    banner("Today's Focus","Berea College · Office of the Provost · Everything you need to tackle today")
    admin_nav_cards(current="today")
    st.markdown("<br>",unsafe_allow_html=True)
    tasks=DB.get_tasks()
    active=[t for t in tasks if t.get("Status","") not in ("Complete","")]
    today_str=date.today().strftime("%Y-%m-%d")
    # Gather all unchecked items from all active tasks.
    all_items=[]  # (task_title, task_id, item_text, item_index, checklist_ref)
    tasks_without_checklist=[]
    for t in active:
        tid=t.get("task_id",""); title=t.get("Title","(untitled)")
        cl_raw=t.get("Checklist","")
        if cl_raw:
            try:
                cl_items=json.loads(cl_raw)
                has_unchecked=False
                for ci,item in enumerate(cl_items):
                    if not item.get("done",False):
                        all_items.append({"task":title,"task_id":tid,"text":item.get("text",""),
                                          "idx":ci,"priority":t.get("Priority",""),"deadline":t.get("Deadline","")})
                        has_unchecked=True
                if not has_unchecked:
                    pass  # all items done for this task
            except: tasks_without_checklist.append(t)
        else:
            tasks_without_checklist.append(t)
    # Sort: urgent first, then by deadline.
    all_items.sort(key=lambda x:("0" if "Urgent" in x["priority"] else "1" if "High" in x["priority"] else "2", x["deadline"]))
    st.caption(f"**{len(all_items)} action items** across **{len(set(i['task_id'] for i in all_items))} tasks** still need your attention. "
               f"Check them off here — progress saves automatically.")
    if not all_items and not tasks_without_checklist:
        st.success("🎉 All checklist items are complete! Check your active tasks for any that are ready to mark as done.")
    # Group by task.
    current_task=None
    for item in all_items:
        if item["task"]!=current_task:
            current_task=item["task"]
            pri_icon=item["priority"].split(" ")[0] if item["priority"] else ""
            st.markdown(f'<div style="margin-top:16px;padding:6px 10px;background:#f0f5fb;border-left:4px solid {BLUE};'
                        f'border-radius:0 6px 6px 0;font-weight:700;font-size:.88rem;color:{BLUE};">'
                        f'{pri_icon} {current_task} <span style="font-weight:400;font-size:.75rem;color:#5a7a96;">'
                        f'· due {item["deadline"]}</span></div>',unsafe_allow_html=True)
        # Interactive checkbox.
        new_val=st.checkbox(item["text"],value=False,key=f"today_{item['task_id']}_{item['idx']}")
        if new_val:
            # User checked this item — update the task's checklist in storage.
            t_data=[t for t in active if t.get("task_id","")==item["task_id"]]
            if t_data:
                try:
                    cl=json.loads(t_data[0].get("Checklist","[]"))
                    cl[item["idx"]]["done"]=True
                    DB.update_task(item["task_id"],"Checklist",json.dumps(cl))
                    st.rerun()
                except: pass
    # Tasks without checklists.
    if tasks_without_checklist:
        st.markdown("<br>",unsafe_allow_html=True)
        sec("📝 Tasks without checklists")
        st.caption("These active tasks don't have a checklist yet. You can generate one from the task detail view.")
        for t in sorted(tasks_without_checklist,key=lambda x:x.get("Deadline","")):
            pri=t.get("Priority",""); pri_icon=pri.split(" ")[0] if pri else ""
            st.markdown(f'- {pri_icon} **{t.get("Title","")}** · {t.get("Status","")} · due {t.get("Deadline","")}')
    # Daily summary button.
    st.markdown("<br>",unsafe_allow_html=True)
    if st.button("📋 Generate today's briefing",key="today_brief"):
        briefing=f"## Today's Focus — {date.today().strftime('%A, %B %d, %Y')}\n\n"
        briefing+=f"**{len(all_items)} items** across **{len(set(i['task_id'] for i in all_items))} tasks** to complete.\n\n"
        cur_t=None
        for item in all_items:
            if item["task"]!=cur_t:
                cur_t=item["task"]
                briefing+=f"\n### {item['priority'].split(' ')[0] if item['priority'] else ''} {cur_t} (due {item['deadline']})\n"
            briefing+=f"- [ ] {item['text']}\n"
        if tasks_without_checklist:
            briefing+="\n### Tasks without checklists\n"
            for t in tasks_without_checklist:
                briefing+=f"- {t.get('Title','')} ({t.get('Status','')}, due {t.get('Deadline','')})\n"
        st.markdown(briefing)
        st.download_button("⬇️ Download as text",data=briefing,file_name=f"focus_{today_str}.md",mime="text/markdown",key="dl_brief")

# ── View: Weekly Planner ─────────────────────────────────────────────────────
def v_planner():
    from datetime import timedelta
    banner("Weekly Planner","Berea College · Office of the Provost · Plan your week at a glance")
    admin_nav_cards(current="planner")
    st.markdown("<br>",unsafe_allow_html=True)
    tasks=DB.get_tasks()
    active=[t for t in tasks if t.get("Status","") not in ("Complete","")]
    today=date.today()
    # Week selector.
    wk_offset=st.session_state.get("planner_week",0)
    pc1,pc2,pc3=st.columns([1,2,1])
    with pc1:
        if st.button("← Previous week",key="pw_prev",use_container_width=True):
            st.session_state.planner_week=wk_offset-1; st.rerun()
    week_start=today-timedelta(days=today.weekday())+timedelta(weeks=wk_offset)
    week_end=week_start+timedelta(days=4)  # Mon-Fri
    with pc2:
        st.markdown(f'<div style="text-align:center;font-size:1.1rem;font-weight:700;color:{BLUE};padding:6px;">'
                    f'{week_start.strftime("%B %d")} — {week_end.strftime("%B %d, %Y")}'
                    f'{"  ·  THIS WEEK" if wk_offset==0 else ""}</div>',unsafe_allow_html=True)
    with pc3:
        if st.button("Next week →",key="pw_next",use_container_width=True):
            st.session_state.planner_week=wk_offset+1; st.rerun()
    if wk_offset!=0:
        if st.button("↩ Back to this week",key="pw_reset"):
            st.session_state.planner_week=0; st.rerun()
    st.markdown("<br>",unsafe_allow_html=True)
    # Build the 5-day grid (Mon-Fri).
    days=[(week_start+timedelta(days=i)) for i in range(5)]
    day_cols=st.columns(5)
    total_items_this_week=0
    for di,d in enumerate(days):
        d_str=d.strftime("%Y-%m-%d")
        is_today=(d==today)
        with day_cols[di]:
            # Day header.
            bg="#e8f1fb" if is_today else "#f5f8fc"
            border=f"border:2px solid {BLUE};" if is_today else "border:1px solid #dde3eb;"
            st.markdown(f'<div style="background:{bg};{border}border-radius:8px;padding:8px 6px;text-align:center;margin-bottom:8px;">'
                        f'<div style="font-weight:700;font-size:.9rem;color:{BLUE};">{d.strftime("%a")}</div>'
                        f'<div style="font-size:.78rem;color:#5a7a96;">{d.strftime("%b %d")}</div>'
                        f'{"<div style=\"font-size:.65rem;font-weight:700;color:#C75B12;\">TODAY</div>" if is_today else ""}'
                        f'</div>',unsafe_allow_html=True)
            # Tasks with this deadline.
            day_tasks=[t for t in active if t.get("Deadline","")==d_str]
            # Also show tasks with checklist items due (we don't have per-item dates, so show tasks whose deadline is this day).
            if day_tasks:
                for t in day_tasks:
                    pri=t.get("Priority",""); pri_icon=pri.split(" ")[0] if pri else ""
                    title=t.get("Title","")
                    short_title=title if len(title)<=20 else title[:18]+"…"
                    stat=t.get("Status",""); sc=S_CSS.get(stat,"s-new")
                    # Mini checklist progress.
                    cl_raw=t.get("Checklist",""); pct_html=""
                    if cl_raw:
                        try:
                            cli=json.loads(cl_raw); ct=len(cli)
                            dn=sum(1 for c in cli if c.get("done",False))
                            pct=int(dn/ct*100) if ct else 0
                            pct_html=(f'<div style="background:#e0e6ed;border-radius:2px;height:4px;margin:4px 0;">'
                                      f'<div style="background:{"#27500A" if pct==100 else BLUE};width:{pct}%;height:4px;border-radius:2px;"></div></div>'
                                      f'<div style="font-size:.6rem;color:#5a7a96;">{dn}/{ct}</div>')
                        except: pass
                    st.markdown(f'<div style="background:#fff;border:1px solid #dde3eb;border-left:3px solid {BLUE};'
                                f'border-radius:0 6px 6px 0;padding:6px 8px;margin-bottom:6px;font-size:.75rem;">'
                                f'<div style="font-weight:700;">{pri_icon} {short_title}</div>'
                                f'<span class="sb {sc}" style="font-size:.6rem;">{stat}</span>'
                                f'{pct_html}</div>',unsafe_allow_html=True)
                    total_items_this_week+=1
            else:
                st.markdown(f'<div style="text-align:center;font-size:.72rem;color:#aaa;padding:12px 0;">—</div>',unsafe_allow_html=True)
    # Overdue tasks (deadline before this week's start).
    overdue=[t for t in active if t.get("Deadline","") and t.get("Deadline","")<week_start.strftime("%Y-%m-%d")]
    if overdue:
        st.markdown("<br>",unsafe_allow_html=True)
        sec("⚠️ Overdue")
        st.caption(f"{len(overdue)} task{'s' if len(overdue)!=1 else ''} with deadlines before this week.")
        for t in sorted(overdue,key=lambda x:x.get("Deadline","")):
            pri=t.get("Priority",""); pi=pri.split(" ")[0] if pri else ""
            st.markdown(f'- {pi} **{t.get("Title","")}** · due {t.get("Deadline","")} · {t.get("Status","")}')
    # Upcoming (next week+).
    upcoming=[t for t in active if t.get("Deadline","")>week_end.strftime("%Y-%m-%d")]
    if upcoming and wk_offset==0:
        with st.expander(f"📅 Coming up after this week ({len(upcoming)} tasks)"):
            for t in sorted(upcoming,key=lambda x:x.get("Deadline",""))[:10]:
                pi=t.get("Priority","").split(" ")[0] if t.get("Priority","") else ""
                st.markdown(f'- {pi} **{t.get("Title","")}** · due {t.get("Deadline","")}')
    st.caption(f"{total_items_this_week} task{'s' if total_items_this_week!=1 else ''} scheduled for this week view.")

# ── View: Delegations (incoming staff task requests + delegate form) ─────────
def v_delegations():
    banner("Task Requests","Berea College · Office of the Provost · Delegate to and from colleagues")
    admin_nav_cards(current="delegations")
    st.markdown("<br>",unsafe_allow_html=True)
    my_email=st.session_state.get("user_email","")
    # ── Incoming delegation requests (pending) ──
    incoming=DB.get_delegations(to_email=my_email,status="Pending")
    sec(f"📥 Incoming task requests ({len(incoming)})")
    if not incoming:
        st.caption("No pending task requests from colleagues right now.")
    else:
        st.caption("Another staff member has asked you to take on these tasks. Accept to add to your task list, or decline.")
        for d in incoming:
            did=d.get("deleg_id","")
            pri=d.get("Priority",""); pi=pri.split(" ")[0] if pri else ""
            with st.expander(f"{pi} {d.get('Title','')} — from {d.get('from_user','')} · due {d.get('Deadline','')}"):
                st.markdown(f'<div style="font-size:.84rem;"><strong>From:</strong> {d.get("from_user","")}<br>'
                            f'<strong>Category:</strong> {d.get("Category","")}<br>'
                            f'<strong>Priority:</strong> {pri}<br>'
                            f'<strong>Deadline:</strong> {d.get("Deadline","")}<br>'
                            f'<strong>Details:</strong> {d.get("Details","(none)")}</div>',unsafe_allow_html=True)
                b1,b2=st.columns(2)
                with b1:
                    if st.button("✅ Accept & add to my tasks",key=f"dgacc_{did}",use_container_width=True):
                        # Create the task assigned to me.
                        DB.save_task({"task_id":gen_task_id(),"created_at":datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "Title":d.get("Title",""),"Category":d.get("Category","Follow-up / admin"),
                            "Priority":pri,"Deadline":d.get("Deadline",""),"Status":"In Progress",
                            "Recurrence":"None — one-time","Details":f"Delegated by {d.get('from_user','')}.\n\n{d.get('Details','')}",
                            "source_ref":"","Visibility":"Shared","Checklist":"",
                            "Calendar_Start":"","Calendar_End":"","Owner":cur_user(),"AssignedTo":cur_user(),"StaffVisible":"Yes"})
                        DB.update_delegation(did,"Status","Accepted")
                        DB.log_audit(did,"delegation_accepted","Pending","Accepted",cur_user(),st.session_state.user_role)
                        st.success("Added to your tasks."); st.rerun()
                with b2:
                    if st.button("🚫 Decline",key=f"dgden_{did}",use_container_width=True):
                        DB.update_delegation(did,"Status","Declined")
                        DB.log_audit(did,"delegation_declined","Pending","Declined",cur_user(),st.session_state.user_role)
                        st.warning("Declined. The sender can see this in their sent requests."); st.rerun()
    # ── Delegate a new task (colleagues + admins only) ──
    if can_delegate():
        st.markdown("<br>",unsafe_allow_html=True)
        sec("📤 Delegate a task to a colleague")
        st.caption("Send a task request to another staff member (admin, colleague, or support). "
                   "They'll accept or decline it — you can't assign to supervisors.")
        # Get eligible staff (admin/colleague/support, not supervisors, not self).
        all_users=DB.get_users()
        staff=[u for u in all_users if u.get("role","") in ("admin","colleague","support")
               and str(u.get("active","True")).lower()!="false"
               and str(u.get("email","")).strip().lower()!=my_email.strip().lower()]
        if not staff:
            st.info("No other staff members are set up yet. Add colleagues or support staff in User Management.")
        else:
            with st.form("delegate_form",clear_on_submit=True):
                pii_banner(compact=True)
                dc1,dc2=st.columns(2)
                with dc1:
                    staff_labels=[f'{u.get("display_name","")} ({u.get("role","")})' for u in staff]
                    chosen=st.selectbox(lbl("Delegate to"),staff_labels)
                with dc2:
                    d_cat=st.selectbox(lbl("Category"),TASK_CATEGORIES)
                d_title=st.text_input(lbl("Task title"),placeholder="e.g., 'Prepare room setup for faculty meeting'")
                dc3,dc4=st.columns(2)
                with dc3: d_pri=st.selectbox(lbl("Priority"),TASK_PRIORITIES,index=2)
                with dc4: d_ddl=st.date_input(lbl("Deadline"),min_value=date.today())
                d_det=st.text_area("Details / context",placeholder="What needs doing, and any helpful context.")
                if st.form_submit_button("📤 Send task request"):
                    if not d_title.strip():
                        st.error("⚠️ Task title is required.")
                    else:
                        chosen_user=staff[staff_labels.index(chosen)]
                        import uuid
                        did=str(uuid.uuid4())[:8]
                        DB.save_delegation({"deleg_id":did,"created_at":datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "from_user":cur_user(),"to_email":chosen_user.get("email",""),
                            "to_name":chosen_user.get("display_name",""),
                            "Title":d_title.strip(),"Category":d_cat,"Priority":d_pri,
                            "Deadline":d_ddl.strftime("%Y-%m-%d"),"Details":d_det.strip(),
                            "Status":"Pending","task_id":""})
                        DB.log_audit(did,"delegation_sent","",chosen_user.get("display_name",""),cur_user(),st.session_state.user_role)
                        st.success(f"Task request sent to {chosen_user.get('display_name','')}. They'll see it under their incoming requests.")
        # ── My sent requests (track status) ──
        all_delegs=DB.get_delegations()
        my_sent=[d for d in all_delegs if d.get("from_user","")==cur_user()]
        if my_sent:
            st.markdown("<br>",unsafe_allow_html=True)
            with st.expander(f"📋 My sent task requests ({len(my_sent)})"):
                for d in sorted(my_sent,key=lambda x:x.get("created_at",""),reverse=True):
                    stat=d.get("Status","Pending")
                    icon={"Pending":"⏳","Accepted":"✅","Declined":"🚫"}.get(stat,"⏳")
                    st.markdown(f'{icon} **{d.get("Title","")}** → {d.get("to_name","")} · {stat} · sent {d.get("created_at","")[:10]}')

# ── View: Task Detail (individual task with all capabilities) ────────────────
def v_task_detail():
    tid=st.session_state.get("detail_task_id","")
    if not tid:
        st.warning("No task selected."); return
    tasks=DB.get_tasks()
    task=[t for t in tasks if t.get("task_id","")==tid]
    if not task:
        st.warning("Task not found."); return
    t=task[0]
    title=t.get("Title","(untitled)")
    banner(f"Task Detail: {title}","Berea College · Office of the Provost")
    if st.button("← Back to My Tasks",key="detail_back"):
        st.session_state.view="tasks"; st.rerun()
    st.markdown("<br>",unsafe_allow_html=True)
    tid_v,pri,uc,stat,vis,rec=_task_props(t)
    _render_task_detail(t,tid_v,stat,pri,uc,rec,vis)
    # ── Calendar download ──
    st.markdown("<br>",unsafe_allow_html=True)
    sec("📅 Add to Calendar")
    ics=generate_ics(t)
    if ics:
        st.download_button("📅 Download .ics for Outlook",data=ics,
                          file_name=f"task_{tid}.ics",mime="text/calendar",key=f"ics_{tid}")
        st.caption("Double-click the downloaded file to add it to Outlook with reminders (1 week + 1 day before).")
    else:
        st.caption("No deadline set — can't generate a calendar event.")
    # ── AI Summary ──
    st.markdown("<br>",unsafe_allow_html=True)
    sec("📝 Task Summary")
    if st.button("Generate summary of what's done and what's still needed",key=f"summ_{tid}"):
        with st.spinner("Generating summary..."):
            summary=generate_task_summary(t)
        st.markdown(summary)
    # ── Generate checklist (if none exists) ──
    cl_raw=t.get("Checklist","")
    if not cl_raw or cl_raw=="[]":
        st.markdown("<br>",unsafe_allow_html=True)
        sec("📋 Generate Checklist")
        st.caption("This task doesn't have a checklist yet. Click below to auto-generate one based on the request details.")
        if st.button("🤖 Auto-generate checklist",key=f"gen_cl_{tid}"):
            with st.spinner("Generating checklist..."):
                # Build a fake submission-like dict for the AI generator.
                subs=DB.get_submissions()
                src=t.get("source_ref","")
                sub_match=[s for s in subs if s.get("ref_number","")==src]
                if sub_match:
                    checklist=generate_ai_checklist(sub_match[0])
                else:
                    # No linked submission — generate from task details.
                    fake_sub={"Request Type":t.get("Category",""),"Your Name":"","Deadline":t.get("Deadline",""),
                              "Additional Notes":t.get("Details",""),"Details JSON":"{}"}
                    checklist=generate_ai_checklist(fake_sub)
                if checklist:
                    DB.update_task(tid,"Checklist",json.dumps(checklist))
                    st.success(f"Generated {len(checklist)} checklist items."); st.rerun()
                else:
                    st.info("Couldn't generate a checklist. You can add items manually, or add an ANTHROPIC_API_KEY to your secrets for AI generation.")

# ── View: User Management (admin only) ───────────────────────────────────────
def v_manage_users():
    banner("User Management","Berea College · Office of the Provost · Add and manage accounts")
    admin_nav_cards(current="users")
    st.markdown("<br>",unsafe_allow_html=True)
    users=DB.get_users()
    active_users=[u for u in users if str(u.get("active","True")).lower()!="false"]
    inactive_users=[u for u in users if str(u.get("active","True")).lower()=="false"]

    sec("📋 Current Users")
    if not active_users:
        st.info("No users in the system yet. Add one below.")
    else:
        for u in sorted(active_users,key=lambda x:x.get("display_name","")):
            role=u.get("role","supervisor"); uid=u.get("user_id","")
            role_icon={"admin":"🔑","colleague":"🤝","support":"🛠️","supervisor":"👤"}.get(role,"👤")
            with st.expander(f'{role_icon} {u.get("display_name","")} — {u.get("email","")} ({role})'):
                st.markdown(f'**Title:** {u.get("title","")}<br>'
                            f'**Role:** {role}<br>'
                            f'**Email:** {u.get("email","")}<br>'
                            f'**Added:** {u.get("created_at","")}',unsafe_allow_html=True)
                nc1,nc2,nc3=st.columns(3)
                with nc1:
                    role_opts=["admin","colleague","support","supervisor"]
                    new_role=st.selectbox("Change role",role_opts,
                        index=role_opts.index(role) if role in role_opts else 3,key=f"ur_{uid}")
                with nc2:
                    new_pw=st.text_input("Set new password",type="password",placeholder="Leave blank to keep current",key=f"up_{uid}")
                with nc3:
                    st.markdown("<br>",unsafe_allow_html=True)
                    if st.button("💾 Save changes",key=f"us_{uid}"):
                        if new_role!=role:
                            DB.update_user(uid,"role",new_role)
                            DB.log_audit(uid,"role_change",role,new_role,cur_user(),st.session_state.user_role)
                        if new_pw.strip():
                            DB.update_user(uid,"password_hash",DB._hash_pw(new_pw.strip()))
                            DB.log_audit(uid,"password_reset","","(admin reset)",cur_user(),st.session_state.user_role)
                        msg=[]
                        if new_role!=role: msg.append(f"role changed to {new_role}")
                        if new_pw.strip(): msg.append("password reset")
                        st.success(f"Updated: {', '.join(msg)}." if msg else "No changes made."); st.rerun()
                if st.button("🚫 Deactivate this user",key=f"ud_{uid}"):
                    DB.update_user(uid,"active","False")
                    DB.log_audit(uid,"deactivated","active","inactive",cur_user(),st.session_state.user_role)
                    st.warning(f'{u.get("display_name","")} has been deactivated.'); st.rerun()

    if inactive_users:
        with st.expander(f"🚫 Deactivated users ({len(inactive_users)})"):
            for u in inactive_users:
                uid=u.get("user_id","")
                st.markdown(f'~~{u.get("display_name","")}~~ — {u.get("email","")} ({u.get("role","")})')
                if st.button(f"Reactivate",key=f"ura_{uid}"):
                    DB.update_user(uid,"active","True")
                    DB.log_audit(uid,"reactivated","inactive","active",cur_user(),st.session_state.user_role)
                    st.success("Reactivated."); st.rerun()

    st.markdown("<br>",unsafe_allow_html=True)
    sec("➕ Add New User")
    with st.form("add_user_form",clear_on_submit=True):
        au1,au2=st.columns(2)
        with au1: au_name=st.text_input(lbl("Full name"),placeholder="e.g., Mary Robert Garrett")
        with au2: au_email=st.text_input(lbl("Email address"),placeholder="e.g., garrettm@berea.edu")
        au3,au4=st.columns(2)
        with au3: au_role=st.selectbox(lbl("Role"),["supervisor","support","colleague","admin"],
                    help="Supervisor = submit requests & view own status. Support = accept/deny requests, manage tasks. Colleague = sees all tasks, delegates to and receives from staff, manages own work (no user management). Admin = full access + user management.")
        with au4: au_title=st.text_input("Title / position",placeholder="e.g., Assoc. Provost & Dean of Faculty")
        au_pw=st.text_input(lbl("Initial password"),type="password",placeholder="Set a temporary password for this user")
        if st.form_submit_button("➕ Add user"):
            errs=[]
            if not au_name.strip(): errs.append("Name is required.")
            if not au_email.strip(): errs.append("Email is required.")
            if not au_pw.strip(): errs.append("Password is required.")
            if DB.get_user_by_email(au_email.strip()):
                errs.append(f"A user with email {au_email.strip()} already exists.")
            if errs:
                for e in errs: st.error(f"⚠️ {e}")
            else:
                import uuid
                new_user={"user_id":str(uuid.uuid4())[:8],
                    "email":au_email.strip().lower(),
                    "display_name":au_name.strip(),"role":au_role,
                    "title":au_title.strip(),
                    "password_hash":DB._hash_pw(au_pw.strip()),
                    "active":"True",
                    "created_at":datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
                DB.save_user(new_user)
                DB.log_audit(new_user["user_id"],"user_created","",au_role,cur_user(),st.session_state.user_role)
                st.success(f"✅ {au_name.strip()} added as {au_role}. They can sign in with their email and password.")
                st.rerun()

# ── View: Notification preferences ───────────────────────────────────────────
def v_prefs():
    user=cur_user(); banner("Settings & Notifications")
    # Determine user's email.
    u_email=st.session_state.get("user_email","")
    u_record=DB.get_user_by_email(u_email) if u_email else None
    info=SUPERVISORS.get(user,{}) if is_sup() else {}
    em=u_email or info.get("email","—") if is_sup() else (u_email or ADMIN_EMAIL or "Not configured")
    st.markdown(f'<div class="ro-notice">📧 Notifications sent to: <strong>{em}</strong></div>',unsafe_allow_html=True)

    # ── Change Password ──────────────────────────────────────────────────────
    sec("🔑 Change Your Password")
    with st.form("pw_change_form",clear_on_submit=True):
        current_pw=st.text_input("Current password",type="password",placeholder="Enter your current password",key="pw_current")
        new_pw1=st.text_input("New password",type="password",placeholder="Enter your new password",key="pw_new1")
        new_pw2=st.text_input("Confirm new password",type="password",placeholder="Re-enter your new password",key="pw_new2")
        if st.form_submit_button("Update password"):
            errs=[]
            if not current_pw.strip(): errs.append("Current password is required.")
            if not new_pw1.strip(): errs.append("New password is required.")
            if new_pw1!=new_pw2: errs.append("New passwords don't match.")
            if len(new_pw1.strip())<6: errs.append("Password must be at least 6 characters.")
            # Verify current password.
            if not errs and u_record:
                if DB._hash_pw(current_pw)!=u_record.get("password_hash",""):
                    errs.append("Current password is incorrect.")
            elif not errs:
                # Fallback: check against secrets for non-sheet users.
                if is_admin() and current_pw!=ADMIN_PASSWORD:
                    errs.append("Current password is incorrect.")
                elif is_sup() and current_pw!=info.get("password",""):
                    errs.append("Current password is incorrect.")
            if errs:
                for e in errs: st.error(f"⚠️ {e}")
            else:
                if u_record:
                    DB.update_user(u_record.get("user_id",""),"password_hash",DB._hash_pw(new_pw1.strip()))
                    DB.log_audit(u_record.get("user_id",""),"password_changed","","(self-service)",user,st.session_state.user_role)
                    st.success("✅ Password updated. Use your new password next time you sign in.")
                else:
                    st.warning("Password change requires a user record in the system. Ask the admin to reset your password.")

    # ── Notification Preferences ─────────────────────────────────────────────
    prefs=DB.get_prefs(user); sec("🔔 Email Notification Settings")
    with st.form("pf_form"):
        if is_sup():
            p1=st.checkbox("Email me when the **status** of one of my requests changes",value=prefs.get("on_status_change",True))
            p2=st.checkbox("Email me when **a note is added** to one of my requests",value=prefs.get("on_new_note",True))
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
    st.markdown("All notifications are sent from the configured Gmail account. They are never sent to external parties and never include confidential request content — only the reference number and a brief description.")

# ── Main router ───────────────────────────────────────────────────────────────
def _check_session_persistence():
    """Check if the user had a recent session (within 10 minutes) and auto-restore it.
    Uses a token stored in URL query params that maps to a session record in Google Sheets."""
    if st.session_state.user: return  # already logged in
    import hashlib, uuid
    params=st.query_params
    token=params.get("session","")
    if not token or not DB.ok: return
    try:
        ws=DB._ws("sessions",["token","user","role","email","mode","timestamp"])
        for r in ws.get_all_records():
            if r.get("token","")==token:
                ts=r.get("timestamp","")
                try:
                    from datetime import datetime as _dt
                    age=(_dt.now()-_dt.strptime(ts,"%Y-%m-%d %H:%M:%S")).total_seconds()
                    if age<600:  # 10 minutes
                        st.session_state.user=r.get("user","")
                        st.session_state.user_role=r.get("role","")
                        st.session_state.user_email=r.get("email","")
                        st.session_state.active_mode=r.get("mode","")
                        st.session_state.view="admin" if r.get("role","") in ("admin","support") else "portal"
                        return
                except: pass
    except: pass

def _save_session_token():
    """Save a session token so the user can reconnect within 10 minutes."""
    if not DB.ok or not st.session_state.user: return
    import hashlib, uuid
    token=hashlib.sha256(f"{st.session_state.user}_{uuid.uuid4()}".encode()).hexdigest()[:16]
    try:
        ws=DB._ws("sessions",["token","user","role","email","mode","timestamp"])
        ws.append_row([token,st.session_state.user,st.session_state.user_role,
                       st.session_state.get("user_email",""),st.session_state.get("active_mode",""),
                       datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
        st.query_params["session"]=token
    except: pass

def main():
    _check_session_persistence()
    if not st.session_state.user: v_login(); return
    # Save/refresh session token on each page load so the 10-minute window resets.
    if "session_saved" not in st.session_state:
        _save_session_token(); st.session_state.session_saved=True
    sidebar(); v=st.session_state.get("view","form")
    if v=="admin" and is_admin_or_support(): v_admin()
    elif v=="tasks" and is_admin_or_support(): v_tasks()
    elif v=="today" and is_admin_or_support(): v_today()
    elif v=="planner" and is_admin_or_support(): v_planner()
    elif v=="delegations" and is_admin_or_support(): v_delegations()
    elif v=="task_detail" and is_admin_or_support(): v_task_detail()
    elif v=="pending" and is_admin_or_support(): view_pending()
    elif v=="urgent" and is_admin_or_support(): view_urgent()
    elif v=="completed" and is_admin_or_support(): v_tasks()
    elif v=="portal" and is_sup(): v_portal()
    elif v=="prefs": v_prefs()
    elif v=="users" and is_admin(): v_manage_users()
    else: v_form()

if __name__=="__main__": main()
