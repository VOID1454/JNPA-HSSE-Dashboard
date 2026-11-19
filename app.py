import io, re
from datetime import date
import pandas as pd
import plotly.express as px
import streamlit as st
from supabase import create_client

st.set_page_config(page_title="JNPA HSSE | Online Dashboard",page_icon="⚓",layout="wide")
st.markdown("<style>.block-container{padding-top:1.2rem}.stMetric{border:1px solid #dce5f0;padding:12px;border-radius:10px}</style>",unsafe_allow_html=True)
try:
 URL=st.secrets["SUPABASE_URL"]
 KEY=st.secrets["SUPABASE_PUBLISHABLE_KEY"]
except (KeyError,FileNotFoundError):
 st.error("Configure SUPABASE_URL and SUPABASE_PUBLISHABLE_KEY in Streamlit secrets. See README.md.")
 st.stop()

def client():
 return create_client(URL,KEY)

if "session" not in st.session_state:
 st.session_state.session=None
st.title("⚓ JNPA | HSSE Training Intelligence")
if st.session_state.session is None:
 st.subheader("Staff login")
 with st.form("login"):
  email=st.text_input("Email")
  password=st.text_input("Password",type="password")
  submitted=st.form_submit_button("Sign in")
 if submitted:
  try:
   result=client().auth.sign_in_with_password({"email":email,"password":password})
   st.session_state.session=result.session
   st.rerun()
  except Exception:
   st.error("Unable to sign in. Check your credentials and account approval.")
 st.caption("Staff accounts are created by the Supabase project administrator. Public sign-up is disabled.")
 st.stop()

def db():
 c=client()
 c.auth.set_session(st.session_state.session.access_token,st.session_state.session.refresh_token)
 return c

try:
 c=db()
 user=c.auth.get_user().user
 profile=c.table("staff_profiles").select("role,display_name,active").eq("user_id",user.id).single().execute().data
 if not profile or not profile["active"]:
  st.error("Account is not approved for dashboard access.")
  st.stop()
except Exception as exc:
 st.session_state.session=None
 st.error("Your session expired or your staff profile is not configured. Please sign in again.")
 st.stop()

role=profile["role"]
can_edit=role in ("admin","data_entry")
with st.sidebar:
 st.write(f"**{profile.get('display_name') or user.email}**")
 st.caption(f"Role: {role.replace('_',' ').title()}")
 if st.button("Sign out"):
  try:c.auth.sign_out()
  except Exception:pass
  st.session_state.session=None
  st.rerun()
 page=st.radio("Module",["Executive dashboard","Organisations","Participants","Training analytics","Training compliance","Data management","Data quality","Staff access"])
 if page=="Staff access" and role!="admin":
  st.warning("Only Admin can view staff access.")
  st.stop()

COLUMNS=["training_date","name","designation","organisation_raw","organisation","phone","email","course","batch","source"]
def fetch_all():
 rows=[]
 start=0
 while True:
  batch=c.table("training").select("*").order("id").range(start,start+999).execute().data or []
  rows.extend(batch)
  if len(batch)<1000:break
  start+=1000
 return pd.DataFrame(rows,columns=["id"]+COLUMNS+["created_at","created_by"])
def csv_button(df,name):
 st.download_button("Download CSV",df.to_csv(index=False).encode("utf-8-sig"),name,mime="text/csv")
def visual(df,key,val,title,kind="bar"):
 if df.empty:
  st.info("No matching data.")
  return
 fig=px.pie(df,names=key,values=val,hole=.38,title=title) if kind=="pie" else px.bar(df,x=key,y=val,title=title)
 st.plotly_chart(fig,use_container_width=True)
def unique(df):
 return df[df.phone.fillna("").astype(str).str.fullmatch(r"\d{10}")].phone.nunique()
def clean(v):return '' if pd.isna(v) else str(v).strip()
def phone(v):
 s=re.sub(r'\D','',clean(v).replace('.0',''))
 if len(s)==12 and s.startswith('91'):s=s[2:]
 if len(s)==11 and s.startswith('0'):s=s[1:]
 return s if len(s)==10 else ''
def org(v):
 s=' '.join(clean(v).upper().replace('.',' ').replace('-',' ').split())
 if s in {'DPW','DP WORLD','NSICT','NSIGT'}:return 'DP World'
 if s in {'BMCT','PSA','PSA MUMBAI'}:return 'PSA Mumbai'
 if s in {'NSFT','NHAVA SHEVA FREEPORT TERMINAL','NHAVA SHEVA FREEPORT TERMINALS'}:return 'NSFT'
 if s in {'NSDT','NHAVA SHEVA DISTRIPARKS TERMINAL','NHAVA SHEVA DISTRIPARKS TREMlNAL'.upper()}:return 'NSDT'
 if s in {'GTI','APM','APM TERMINALS','MAERSK','MEARSK'}:return 'APM Terminals'
 if 'POLICE' in s:return 'Maharashtra Police'
 return clean(v).strip() or 'Not specified'

def read_excel(blob):
 x=pd.ExcelFile(io.BytesIO(blob)); frames=[]
 for sheet in x.sheet_names:
  df=pd.read_excel(io.BytesIO(blob),sheet_name=sheet,header=None)
  hits=df.index[df.astype(str).apply(lambda r: r.str.strip().str.upper().eq('NAME').any(),axis=1)].tolist()
  if not hits:continue
  h=hits[0]; headers=[clean(v).upper() for v in df.iloc[h].tolist()];df=df.iloc[h+1:].copy();df.columns=headers
  def find(*names):
   for n in names:
    if n in df.columns:return df[n]
   return pd.Series(['']*len(df),index=df.index)
  out=pd.DataFrame({'training_date':find('DATE'),'name':find('NAME','PARTICIPANT NAME'),'designation':find('DESIGNATION'),'organisation_raw':find('ORGANISATION','ORGANIZATION'),'organisation':find('STANDARDISED ORGANISATION','STANDARDIZED ORGANISATION'),'phone':find('CONTACT NO.','CONTACT NO','MOBILE NUMBER','PHONE'),'email':find('EMAIL','EMAIL ID'),'course':find('COURSE','TRAINING','TRAINING TOPIC'),'batch':find('BATCH')})
  out=out[out.name.notna() & out.name.astype(str).str.strip().ne('')]
  frames.append(out)
 if not frames:raise ValueError('No sheet with a NAME column was found.')
 return pd.concat(frames,ignore_index=True)

def prepare(df,source):
 df=df.copy()
 for col in ['name','designation','organisation_raw','organisation','phone','email','course','batch']:
  df[col]=df[col].map(clean)
 df['phone']=df.phone.map(phone)
 df['course']=df['course'].replace('', 'Not Specified')
 df['organisation']=[org(a if a else b) for a,b in zip(df.organisation,df.organisation_raw)]
 # Parse dates individually to support mixed timezone-aware and naive Excel values.
 def normalise_training_date(value):
  if pd.isna(value) or str(value).strip()=='':return ''
  try:
   parsed=pd.to_datetime(value,errors='coerce')
   return '' if pd.isna(parsed) else parsed.strftime('%Y-%m-%d')
  except (ValueError,TypeError,OverflowError):return ''
 df['training_date']=df['training_date'].map(normalise_training_date)
 df['source']=source
 return df[['training_date','name','designation','organisation_raw','organisation','phone','email','course','batch','source']]


d=fetch_all().fillna("")
with st.sidebar:
 st.divider()
 orgs=st.multiselect("Organisation",sorted(d.organisation.unique().tolist()) if len(d) else [])
 courses=st.multiselect("Training course",sorted(d.course.unique().tolist()) if len(d) else [])
 search=st.text_input("Name / contact / email")
 start=st.date_input("From date",value=None)
 end=st.date_input("To date",value=None)
f=d.copy()
if orgs:f=f[f.organisation.isin(orgs)]
if courses:f=f[f.course.isin(courses)]
if search:
 f=f[f[["name","phone","email"]].astype(str).apply(lambda col:col.str.contains(re.escape(search),case=False)).any(axis=1)]
if start:f=f[f.training_date.astype(str)>=str(start)]
if end:f=f[f.training_date.astype(str)<=str(end)]

if page=="Executive dashboard":
 a,b,c1,e,g=st.columns(5)
 a.metric("Training attendances",f"{len(f):,}")
 b.metric("Unique valid contacts",f"{unique(f):,}")
 c1.metric("Organisations",f.organisation.nunique())
 e.metric("Courses",f.course.nunique())
 g.metric("Missing / invalid phones",int((~f.phone.astype(str).str.fullmatch(r"\d{10}")).sum()))
 st.caption("Attendances are records; unique participants are distinct valid contact numbers.")
 summary=f.groupby("organisation").size().reset_index(name="Attendances").sort_values("Attendances",ascending=False)
 l,r=st.columns(2)
 with l:visual(summary.head(15),"organisation","Attendances","Top organisations")
 with r:visual(summary.head(10),"organisation","Attendances","Organisation share","pie")
 trend=f.copy()
 trend["month"]=trend.training_date.astype(str).str.slice(0,7)
 trend=trend[trend.month.str.fullmatch(r"\\d{4}-\\d{2}")].groupby("month").size().reset_index(name="Attendances")
 if len(trend):st.plotly_chart(px.line(trend,x="month",y="Attendances",markers=True,title="Monthly training trend"),use_container_width=True)
elif page=="Organisations":
 a=f.groupby("organisation").agg(attendances=("id","size"),courses=("course","nunique")).reset_index()
 u=f[f.phone.astype(str).str.fullmatch(r"\d{10}")].groupby("organisation").phone.nunique().rename("unique_participants")
 a=a.join(u,on="organisation").fillna({"unique_participants":0}).sort_values("attendances",ascending=False)
 st.dataframe(a,use_container_width=True,hide_index=True);csv_button(a,"organisation_summary.csv")
 visual(a.head(20),"organisation","attendances","Organisation attendances")
elif page=="Participants":
 good=f[f.phone.astype(str).str.fullmatch(r"\d{10}")]
 people=good.sort_values("training_date").groupby("phone").agg(name=("name","last"),organisation=("organisation","last"),attendances=("id","size"),courses=("course","nunique"),last_training=("training_date","max")).reset_index()
 st.dataframe(people,use_container_width=True,hide_index=True);csv_button(people,"participant_directory.csv")
 number=st.text_input("Search complete training history by 10-digit contact number")
 if number:
  history=d[d.phone==phone(number)].sort_values("training_date",ascending=False)
  st.dataframe(history[["training_date","name","organisation","course","batch","designation","phone"]],use_container_width=True,hide_index=True)
  csv_button(history,"participant_history.csv")
elif page=="Training analytics":
 a=f.groupby("course").size().reset_index(name="Attendances").sort_values("Attendances",ascending=False)
 st.dataframe(a,use_container_width=True,hide_index=True)
 l,r=st.columns(2)
 with l:visual(a.head(15),"course","Attendances","Top training courses")
 with r:visual(a.head(10),"course","Attendances","Course share","pie")
 matrix=pd.crosstab(f.organisation,f.course).reset_index()
 st.dataframe(matrix,use_container_width=True);csv_button(matrix,"organisation_course_matrix.csv")
elif page=="Training compliance":
 st.subheader("Training Compliance Dashboard")
 st.caption("Compliance is based on Admin-entered organisation/course targets, not inferred from attendance. Completed means distinct valid 10-digit contacts with at least one matching training record.")
 try:
  target_rows=c.table("compliance_targets").select("*").order("organisation").range(0,999).execute().data or []
 except Exception:
  st.error("Compliance database table is not installed. Run compliance_setup.sql in Supabase SQL Editor first.")
  st.stop()
 targets=pd.DataFrame(target_rows)
 st.markdown("#### Compliance targets")
 if targets.empty:
  st.info("No targets configured yet. An Admin must enter required participant counts for each organisation and training course before compliance percentages or pending counts can be calculated.")
 else:
  base=d.copy()
  base["phone"]=base.phone.astype(str).str.strip()
  base=base[base.phone.str.fullmatch(r"\\d{10}")]
  counts=base.groupby(["organisation","course"]).phone.nunique().rename("completed_unique").reset_index()
  comp=targets.merge(counts,on=["organisation","course"],how="left")
  comp["completed_unique"]=comp.completed_unique.fillna(0).astype(int)
  comp["required_participants"]=pd.to_numeric(comp.required_participants,errors="coerce").fillna(0).astype(int)
  comp["credited"]=comp[["completed_unique","required_participants"]].min(axis=1)
  comp["pending"]= (comp.required_participants-comp.credited).clip(lower=0)
  comp["compliance_pct"]=(100*comp.credited/comp.required_participants.where(comp.required_participants>0)).round(1)
  comp["extra_attendances_above_target"]=(comp.completed_unique-comp.required_participants).clip(lower=0)
  filtered=comp.copy()
  if orgs:filtered=filtered[filtered.organisation.isin(orgs)]
  if courses:filtered=filtered[filtered.course.isin(courses)]
  # Date/name filters apply to training attendance views, but not the all-time
  # compliance baseline, so we do not silently alter target denominators.
  if search or start or end:
   st.info("Compliance uses all recorded training history. Sidebar name/date filters do not change target completion; use organisation and course filters here.")
  required=int(filtered.required_participants.sum())
  credited=int(filtered.credited.sum())
  pending=int(filtered.pending.sum())
  pct=round(100*credited/required,1) if required else 0
  a,b,cc,dd=st.columns(4)
  a.metric("Required participation slots",f"{required:,}")
  b.metric("Completed slots",f"{credited:,}")
  cc.metric("Pending slots",f"{pending:,}")
  dd.metric("Compliance",f"{pct}%")
  st.caption("These are organisation-course participation slots, not unique people across all courses. Extra attendances do not exceed 100% credit.")
  display=filtered[["organisation","course","required_participants","completed_unique","credited","pending","compliance_pct","extra_attendances_above_target"]].sort_values(["organisation","course"])
  st.dataframe(display,use_container_width=True,hide_index=True)
  csv_button(display,"jnpa_compliance_by_organisation_course.csv")
  if len(filtered):
   orgsum=filtered.groupby("organisation",as_index=False)[["required_participants","credited","pending"]].sum()
   orgsum["compliance_pct"]=(100*orgsum.credited/orgsum.required_participants.where(orgsum.required_participants>0)).round(1)
   l,r=st.columns(2)
   with l:visual(orgsum,"organisation","pending","Pending participation slots by organisation")
   with r:visual(orgsum,"organisation","compliance_pct","Compliance percentage by organisation")
   st.subheader("Organisation summary")
   st.dataframe(orgsum,use_container_width=True,hide_index=True)
   csv_button(orgsum,"jnpa_compliance_organisation_summary.csv")
 st.divider()
 if role=="admin":
  st.subheader("Admin: Set or update compliance targets")
  st.caption("Enter the official required number of distinct participants for a specific organisation and course. Updating an existing pair replaces its target, not attendance records.")
  known_orgs=sorted(set(d.organisation.astype(str)) | (set(targets.organisation.astype(str)) if not targets.empty else set()))
  known_courses=sorted(set(d.course.astype(str)) | (set(targets.course.astype(str)) if not targets.empty else set()))
  with st.form("target_editor"):
   selected_org=st.selectbox("Organisation",known_orgs) if known_orgs else st.text_input("Organisation")
   selected_course=st.selectbox("Training course",known_courses) if known_courses else st.text_input("Training course")
   required_n=st.number_input("Required participants",min_value=1,step=1,value=1)
   if st.form_submit_button("Save target"):
    if not selected_org or not selected_course:st.error("Organisation and course are required.")
    else:
     try:
      c.table("compliance_targets").upsert({"organisation":selected_org,"course":selected_course,"required_participants":int(required_n),"updated_by":user.id},on_conflict="organisation,course").execute()
      st.success("Compliance target saved.");st.rerun()
     except Exception as exc:st.error(f"Unable to save target: {exc}")
  if not targets.empty:
   with st.expander("Remove an incorrect target"):
    choices={f"{r['organisation']} — {r['course']} (target {r['required_participants']})":r["id"] for r in target_rows}
    chosen=st.selectbox("Target to remove",list(choices))
    confirm=st.checkbox("I confirm this target should be removed")
    if st.button("Delete target",disabled=not confirm):
     try:c.table("compliance_targets").delete().eq("id",choices[chosen]).execute();st.rerun()
     except Exception as exc:st.error(str(exc))
 else:
  st.info("Only Admin can change compliance targets. Data Entry and Management can view and export compliance reports.")
 st.caption("Individual pending-person lists require an official eligible/nominated participant roster. Attendance data alone cannot identify who was required to attend but did not.")
elif page=="Data quality":
 invalid=f[~f.phone.astype(str).str.fullmatch(r"\d{10}")]
 st.metric("Attendance records with missing/invalid contact",len(invalid))
 st.dataframe(invalid,use_container_width=True,hide_index=True)
 names=f[f.phone.astype(str).str.fullmatch(r"\d{10}")].groupby("phone").name.nunique()
 st.subheader("Different names sharing one contact")
 st.dataframe(f[f.phone.isin(names[names>1].index)][["phone","name","organisation"]].drop_duplicates(),use_container_width=True,hide_index=True)
elif page=="Staff access":
 st.subheader("Staff access and roles")
 st.info("Create users in Supabase Authentication > Users, then approve their user UUID in the staff_profiles table. Only approved users can sign in.")
 profiles=c.table("staff_profiles").select("user_id,display_name,role,active").execute().data
 st.dataframe(profiles,use_container_width=True,hide_index=True)
 st.caption("To change staff permissions, update staff_profiles in the Supabase SQL Editor or Table Editor. Do not enable public sign-up.")
elif page=="Data management":
 if not can_edit:
  st.info("Your Management role has read-only access.")
 else:
  st.subheader("Add training attendance")
  with st.form("new"):
   a,b,c1=st.columns(3)
   with a:dt=st.date_input("Training date",date.today());name=st.text_input("Participant name");mobile=st.text_input("Contact number")
   with b:organisation=st.text_input("Organisation");course=st.text_input("Training course");batch=st.text_input("Batch")
   with c1:designation=st.text_input("Designation");email=st.text_input("Email")
   if st.form_submit_button("Save attendance"):
    if not name.strip() or not course.strip():st.error("Name and course are required.")
    else:
     row=dict(training_date=str(dt),name=name.strip(),phone=phone(mobile),organisation_raw=organisation,organisation=org(organisation),course=course.strip(),batch=batch,designation=designation,email=email,source="manual",created_by=user.id)
     try:c.table("training").insert(row).execute();st.success("Attendance saved.");st.rerun()
     except Exception as exc:st.error(f"Could not save: {exc}")
  st.subheader("Bulk Excel import")
  up=st.file_uploader("Upload Excel (.xlsx)",type=["xlsx"])
  if up:
   try:
    incoming=prepare(read_excel(up.getvalue()),up.name)
    st.caption(f"{len(incoming):,} rows detected. Existing records are not replaced.")
    st.dataframe(incoming.head(10),use_container_width=True)
    if st.button("Append new attendance records"):
     # Compare business fields, not database IDs; do not overwrite existing records.
     keys=["training_date","name","designation","organisation_raw","organisation","phone","email","course","batch"]
     def sig(row):return tuple(str(row.get(k,"") or "") for k in keys)
     existing={sig(r) for r in d.to_dict("records")}
     added=0
     for row in incoming.to_dict("records"):
      cleanrow={k:(None if pd.isna(v) or v=="" else str(v)) for k,v in row.items()}
      cleanrow['course']=cleanrow.get('course') or 'Not Specified'
      fingerprint=sig(cleanrow)
      if fingerprint in existing:continue
      cleanrow["created_by"]=user.id
      c.table("training").insert(cleanrow).execute()
      existing.add(fingerprint);added+=1
     st.success(f"Imported {added:,} new rows; skipped {len(incoming)-added:,} duplicates.");st.rerun()
   except Exception as exc:st.error(f"Import failed: {exc}")
  st.subheader("Edit / delete a record")
  st.dataframe(f[["id","training_date","name","phone","organisation","course"]].tail(300),use_container_width=True,hide_index=True)
  ids=f.id.tolist()
  if ids:
   rid=st.selectbox("Select record ID",ids)
   rec=d[d.id==rid].iloc[0]
   with st.form("edit"):
    name2=st.text_input("Name",rec["name"])
    phone2=st.text_input("Phone",rec["phone"])
    org2=st.text_input("Organisation",rec["organisation_raw"] or rec["organisation"])
    course2=st.text_input("Course",rec["course"])
    dt2=st.text_input("Training date YYYY-MM-DD",rec["training_date"])
    if st.form_submit_button("Save changes"):
     try:
      c.table("training").update(dict(name=name2,phone=phone(phone2),organisation_raw=org2,organisation=org(org2),course=course2,training_date=dt2)).eq("id",int(rid)).execute()
      st.success("Updated");st.rerun()
     except Exception as exc:st.error(str(exc))
   if st.checkbox("I confirm deletion of the selected record") and st.button("Delete selected record"):
    try:c.table("training").delete().eq("id",int(rid)).execute();st.rerun()
    except Exception as exc:st.error(str(exc))
 st.subheader("Export")
 csv_button(d,"all_training_records.csv")
