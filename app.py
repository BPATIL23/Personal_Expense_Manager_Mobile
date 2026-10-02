from flask import Flask, request, redirect, url_for, send_file, render_template_string, flash, jsonify
import sqlite3, os
from datetime import date
from openpyxl import Workbook

try:
    import psycopg2
    from psycopg2.extras import RealDictCursor
except ImportError:
    psycopg2 = None
    RealDictCursor = None

BASE=os.path.dirname(os.path.abspath(__file__))
DB=os.path.join(BASE,"expenses.db")
DATABASE_URL=os.environ.get("DATABASE_URL","").strip()
USE_POSTGRES=bool(DATABASE_URL)
app=Flask(__name__)
app.secret_key=os.environ.get("SECRET_KEY","expense-manager")

def db():
    if USE_POSTGRES:
        if psycopg2 is None:
            raise RuntimeError("psycopg2-binary is required when DATABASE_URL is set")
        dsn=DATABASE_URL.replace("postgres://","postgresql://",1)
        return psycopg2.connect(dsn, sslmode="require", cursor_factory=RealDictCursor)
    c=sqlite3.connect(DB)
    c.row_factory=sqlite3.Row
    return c

def execute(c, sql, params=()):
    return c.execute(sql.replace("?", "%s") if USE_POSTGRES else sql, params)

def init():
    c=db()
    if USE_POSTGRES:
        execute(c, "CREATE TABLE IF NOT EXISTS transactions (id BIGSERIAL PRIMARY KEY,type TEXT,tx_date TEXT,category TEXT,description TEXT,amount DOUBLE PRECISION,payment_mode TEXT,remarks TEXT)")
        execute(c, "CREATE TABLE IF NOT EXISTS members (id BIGSERIAL PRIMARY KEY,name TEXT UNIQUE NOT NULL,mobile TEXT,remarks TEXT)")
        execute(c, "CREATE TABLE IF NOT EXISTS person_ledger (id BIGSERIAL PRIMARY KEY,person TEXT,tx_date TEXT,particular TEXT,paid DOUBLE PRECISION DEFAULT 0,returned DOUBLE PRECISION DEFAULT 0,remarks TEXT)")
    else:
        execute(c, "CREATE TABLE IF NOT EXISTS transactions (id INTEGER PRIMARY KEY AUTOINCREMENT,type TEXT,tx_date TEXT,category TEXT,description TEXT,amount REAL,payment_mode TEXT,remarks TEXT)")
        execute(c, "CREATE TABLE IF NOT EXISTS members (id INTEGER PRIMARY KEY AUTOINCREMENT,name TEXT UNIQUE NOT NULL,mobile TEXT,remarks TEXT)")
        execute(c, "CREATE TABLE IF NOT EXISTS person_ledger (id INTEGER PRIMARY KEY AUTOINCREMENT,person TEXT,tx_date TEXT,particular TEXT,paid REAL DEFAULT 0,returned REAL DEFAULT 0,remarks TEXT)")
    c.commit(); c.close()

CSS="""
body{margin:0;font-family:Arial;background:#f5f7fb;color:#1f2937}
header{background:#111827;color:white;padding:15px 22px;display:flex;justify-content:space-between;align-items:center;gap:15px}
.brand{font-size:21px;font-weight:bold}.nav{display:flex;gap:6px;flex-wrap:wrap}.nav a{color:white;text-decoration:none;padding:9px 12px;border-radius:7px}.nav a:hover{background:#2563eb}
main{max-width:1200px;margin:25px auto;padding:0 16px}.title{display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap}
h1{margin:0 0 5px}.card{background:white;border:1px solid #e5e7eb;border-radius:12px;padding:18px;margin:15px 0}
.cards{display:grid;grid-template-columns:repeat(4,1fr);gap:14px}.card .label{color:#6b7280;font-size:13px}.num{font-size:25px;font-weight:bold;margin-top:8px}
.grid{display:grid;grid-template-columns:repeat(2,1fr);gap:15px}label{font-weight:bold;font-size:13px;display:block;margin-bottom:6px}
input,select,textarea{width:100%;padding:11px;border:1px solid #d1d5db;border-radius:8px;box-sizing:border-box}textarea{min-height:80px}
.btn{display:inline-block;border:0;background:#2563eb;color:white;padding:10px 14px;border-radius:8px;text-decoration:none;cursor:pointer}.green{background:#15803d}.red{background:#dc2626}.gray{background:#374151}.light{background:#e5e7eb;color:#111827}
.actions{display:flex;gap:8px;flex-wrap:wrap}.table{overflow:auto;background:white;border:1px solid #e5e7eb;border-radius:12px}
table{width:100%;border-collapse:collapse;min-width:760px}th,td{padding:11px;border-bottom:1px solid #eee;text-align:left}th{background:#f1f5f9}.right{text-align:right}
.people{display:grid;grid-template-columns:repeat(3,1fr);gap:14px}.person{background:white;border:1px solid #e5e7eb;border-radius:12px;padding:17px}
.notice{background:#eff6ff;border:1px solid #bfdbfe;padding:12px;border-radius:8px}.flash{background:#fef3c7;padding:12px;border-radius:8px;margin:10px 0}
@media(max-width:800px){.cards,.people,.grid{grid-template-columns:1fr 1fr}}
@media(max-width:550px){.cards,.people,.grid{grid-template-columns:1fr}header{position:sticky;top:0;z-index:20;flex-direction:column;align-items:stretch;padding:12px}.brand{font-size:18px}.nav{display:grid;grid-template-columns:repeat(3,1fr);gap:5px}.nav a{text-align:center;padding:10px 3px;font-size:12px;background:#1f2937}main{margin:12px auto;padding:0 10px}.card{padding:14px;margin:10px 0}.btn{min-height:42px;box-sizing:border-box;display:inline-flex;align-items:center;justify-content:center}input,select,textarea{font-size:16px;min-height:44px}table{min-width:760px}th,td{padding:9px;font-size:13px}h1{font-size:24px}.num{font-size:22px}}
"""

def page(title,body):
    nav="""<header><div class='brand'>Personal Expense Manager</div>
    <nav class='nav'><a href='/'>Dashboard</a><a href='/add'>Add Expense</a><a href='/transactions'>Transactions</a>
    <a href='/report'>Reports</a><a href='/people'>Member Ledger</a></nav></header>"""
    return render_template_string("""<!doctype html><html><head><meta charset='utf-8'><meta name='viewport' content='width=device-width,initial-scale=1,maximum-scale=1,user-scalable=no'><link rel='manifest' href='/manifest.webmanifest'><meta name='theme-color' content='#111827'><meta name='mobile-web-app-capable' content='yes'><meta name='apple-mobile-web-app-capable' content='yes'><title>{{title}}</title><style>{{css|safe}}</style></head><body>{{nav|safe}}<main>{{body|safe}}</main><script>if('serviceWorker' in navigator){navigator.serviceWorker.register('/sw.js').catch(()=>{});}</script></body></html>""",title=title,css=CSS,nav=nav,body=body)

@app.route("/")
def home():
    c=db()
    income=execute(c, "SELECT COALESCE(SUM(amount),0) x FROM transactions WHERE type='Income'").fetchone()['x']
    expense=execute(c, "SELECT COALESCE(SUM(amount),0) x FROM transactions WHERE type='Expense'").fetchone()['x']
    members=execute(c, "SELECT COUNT(*) x FROM members").fetchone()['x']
    paid=execute(c, "SELECT COALESCE(SUM(paid),0) x FROM person_ledger").fetchone()['x']
    returned=execute(c, "SELECT COALESCE(SUM(returned),0) x FROM person_ledger").fetchone()['x']
    c.close()
    body=render_template_string("""<div class='title'><div><h1>Dashboard</h1><small>Expense and Member Ledger</small></div>
    <div class='actions'><a class='btn' href='/add'>+ Add Expense</a><a class='btn green' href='/member/add'>+ Add Member</a></div></div>
    <div class='cards'>
    <div class='card'><div class='label'>Total Income</div><div class='num'>₹ {{'%.2f'|format(income)}}</div></div>
    <div class='card'><div class='label'>Total Expense</div><div class='num'>₹ {{'%.2f'|format(expense)}}</div></div>
    <div class='card'><div class='label'>Available Balance</div><div class='num'>₹ {{'%.2f'|format(income-expense)}}</div></div>
    <div class='card'><div class='label'>Members</div><div class='num'>{{members}}</div></div></div>
    <div class='card'><h2>Member Ledger Summary</h2><p>Total Paid: <b>₹ {{'%.2f'|format(paid)}}</b></p><p>Total Return: <b>₹ {{'%.2f'|format(returned)}}</b></p><p>Balance to Receive: <b>₹ {{'%.2f'|format(paid-returned)}}</b></p>
    <div class='actions'><a class='btn green' href='/people'>Open Member Ledger</a><a class='btn gray' href='/ledger/add'>+ Paid / Return</a></div></div>""",income=income,expense=expense,members=members,paid=paid,returned=returned)
    return page("Dashboard",body)

@app.route("/add",methods=["GET","POST"])
def add():
    cats=["Food","Travel","Petrol","Shopping","Medical","Electricity","Mobile/Internet","Home","Entertainment","Education","Other"]
    modes=["Cash","Bank","UPI","Card","Other"]
    if request.method=="POST":
        c=db();execute(c, "INSERT INTO transactions(type,tx_date,category,description,amount,payment_mode,remarks) VALUES(?,?,?,?,?,?,?)",
        (request.form['type'],request.form['tx_date'],request.form.get('category',''),request.form.get('description',''),float(request.form['amount']),request.form.get('payment_mode',''),request.form.get('remarks','')))
        c.commit();c.close();return redirect('/transactions')
    body=render_template_string("""<div class='title'><h1>Add Expense / Income</h1><a class='btn light' href='/'>Back</a></div><div class='card'><form method='post'><div class='grid'>
    <div><label>Type</label><select name='type'><option>Expense</option><option>Income</option></select></div>
    <div><label>Date</label><input type='date' name='tx_date' value='{{today}}' required></div>
    <div><label>Category</label><select name='category'>{%for x in cats%}<option>{{x}}</option>{%endfor%}</select></div>
    <div><label>Amount ₹</label><input type='number' name='amount' step='.01' min='0' required></div>
    <div><label>Payment Mode</label><select name='payment_mode'>{%for x in modes%}<option>{{x}}</option>{%endfor%}</select></div>
    <div><label>Description</label><input name='description'></div>
    <div style='grid-column:1/-1'><label>Remarks</label><textarea name='remarks'></textarea></div></div><br><button class='btn'>Save</button></form></div>""",cats=cats,modes=modes,today=date.today().isoformat())
    return page("Add",body)

@app.route("/transactions")
def transactions():
    c=db();rows=execute(c, "SELECT * FROM transactions ORDER BY tx_date DESC,id DESC").fetchall();c.close()
    body=render_template_string("""<div class='title'><h1>Transactions</h1><a class='btn' href='/add'>+ Add Entry</a></div><div class='table'><table><tr><th>Date</th><th>Type</th><th>Category</th><th>Description</th><th>Mode</th><th class='right'>Amount</th></tr>
    {%for r in rows%}<tr><td>{{r.tx_date}}</td><td>{{r.type}}</td><td>{{r.category}}</td><td>{{r.description}}</td><td>{{r.payment_mode}}</td><td class='right'>₹ {{'%.2f'|format(r.amount)}}</td></tr>{%else%}<tr><td colspan='6'>No entries.</td></tr>{%endfor%}</table></div>""",rows=rows)
    return page("Transactions",body)

@app.route("/report")
def report():
    c=db();rows=execute(c, """SELECT substr(tx_date,1,7) month,
    SUM(CASE WHEN type='Income' THEN amount ELSE 0 END) income,
    SUM(CASE WHEN type='Expense' THEN amount ELSE 0 END) expense
    FROM transactions GROUP BY substr(tx_date,1,7) ORDER BY month DESC""").fetchall();c.close()
    body=render_template_string("""<h1>Reports</h1><div class='card'><div class='table'><table><tr><th>Month</th><th>Income</th><th>Expense</th><th>Balance</th></tr>
    {%for r in rows%}<tr><td>{{r.month}}</td><td>₹ {{'%.2f'|format(r.income)}}</td><td>₹ {{'%.2f'|format(r.expense)}}</td><td>₹ {{'%.2f'|format(r.income-r.expense)}}</td></tr>{%endfor%}</table></div></div>""",rows=rows)
    return page("Reports",body)

# ---------- MEMBER MASTER ----------
@app.route("/member/add",methods=["GET","POST"])
def member_add():
    if request.method=="POST":
        name=request.form.get("name","").strip();mobile=request.form.get("mobile","").strip();remarks=request.form.get("remarks","").strip()
        try:
            c=db();execute(c, "INSERT INTO members(name,mobile,remarks) VALUES(?,?,?)",(name,mobile,remarks));c.commit();c.close();return redirect('/people')
        except Exception: flash("Member already exists.")
    body=render_template_string("""<div class='title'><h1>➕ Add Member</h1><a class='btn light' href='/people'>Back</a></div>
    {%with m=get_flashed_messages()%}{%for x in m%}<div class='flash'>{{x}}</div>{%endfor%}{%endwith%}
    <div class='card'><form method='post'><div class='grid'><div><label>Member Name *</label><input name='name' required placeholder='Kavya Maam'></div>
    <div><label>Mobile No. (Optional)</label><input name='mobile'></div><div style='grid-column:1/-1'><label>Remarks</label><textarea name='remarks'></textarea></div></div><br><button class='btn green'>Save Member</button></form></div>""")
    return page("Add Member",body)

@app.route("/people")
def people():
    c=db()
    members=execute(c, """SELECT m.*,COALESCE(SUM(l.paid),0) paid,COALESCE(SUM(l.returned),0) returned,
    COALESCE(SUM(l.paid-l.returned),0) balance FROM members m LEFT JOIN person_ledger l ON l.person=m.name
    GROUP BY m.id ORDER BY m.name""").fetchall()
    old=execute(c, """SELECT person,SUM(paid)paid,SUM(returned)returned,SUM(paid-returned)balance
    FROM person_ledger WHERE person NOT IN (SELECT name FROM members) GROUP BY person""").fetchall()
    c.close()
    body=render_template_string("""<div class='title'><div><h1>Member Ledger</h1><small>Add member once, then select name in Paid / Return.</small></div>
    <div class='actions'><a class='btn green' href='/member/add'>+ Add Member</a><a class='btn' href='/ledger/add'>+ Paid / Return</a><a class='btn light' href='/ledger/export'>Export Excel</a></div></div>
    <div class='notice'><b>Paid</b> = आपण दिले &nbsp; | &nbsp; <b>Return</b> = परत मिळाले &nbsp; | &nbsp; <b>Balance</b> = अजून परत मिळायचे</div>
    <div class='people'>{%for p in members%}<div class='person'><h2>{{p.name}}</h2><small>{{p.mobile or ''}}</small><p>Paid: <b>₹ {{'%.2f'|format(p.paid)}}</b></p><p>Return: <b>₹ {{'%.2f'|format(p.returned)}}</b></p><p>Balance: <b>₹ {{'%.2f'|format(p.balance)}}</b></p><a class='btn' href='/ledger?person={{p.name|urlencode}}'>View Ledger</a></div>{%endfor%}
    {%for p in old%}<div class='person'><h2>{{p.person}}</h2><small>Old ledger member</small><p>Paid: <b>₹ {{'%.2f'|format(p.paid)}}</b></p><p>Return: <b>₹ {{'%.2f'|format(p.returned)}}</b></p><p>Balance: <b>₹ {{'%.2f'|format(p.balance)}}</b></p><a class='btn' href='/ledger?person={{p.person|urlencode}}'>View Ledger</a></div>{%endfor%}
    {%if not members and not old%}<div class='card'>No members. Click + Add Member.</div>{%endif%}</div>""",members=members,old=old)
    return page("Member Ledger",body)

@app.route("/ledger/add",methods=["GET","POST"])
def ledger_add():
    c=db();members=execute(c, "SELECT name FROM members ORDER BY name").fetchall();c.close()
    if request.method=="POST":
        person=request.form.get("person","");paid=float(request.form.get("paid","0") or 0);ret=float(request.form.get("returned","0") or 0)
        if paid>0 and ret>0: flash("Enter either Paid OR Return.")
        elif paid<=0 and ret<=0: flash("Enter Paid or Return amount.")
        else:
            c=db();execute(c, "INSERT INTO person_ledger(person,tx_date,particular,paid,returned,remarks) VALUES(?,?,?,?,?,?)",
            (person,request.form['tx_date'],request.form.get('particular',''),paid,ret,request.form.get('remarks','')));c.commit();c.close();return redirect('/ledger?person='+person)
    body=render_template_string("""<div class='title'><h1>💵 Paid / Return</h1><div class='actions'><a class='btn green' href='/member/add'>+ Add Member</a><a class='btn light' href='/people'>Back</a></div></div>
    {%with m=get_flashed_messages()%}{%for x in m%}<div class='flash'>{{x}}</div>{%endfor%}{%endwith%}
    <div class='card'><form method='post'><div class='grid'><div><label>Member *</label><select name='person' required><option value=''>-- Select Member --</option>{%for m in members%}<option>{{m.name}}</option>{%endfor%}</select></div>
    <div><label>Date</label><input type='date' name='tx_date' value='{{today}}' required></div><div><label>Paid - आपण दिले ₹</label><input type='number' name='paid' value='0' min='0' step='.01'></div>
    <div><label>Return - परत मिळाले ₹</label><input type='number' name='returned' value='0' min='0' step='.01'></div>
    <div style='grid-column:1/-1'><label>Particular</label><input name='particular' placeholder='Trip Cash Given / Cash Returned'></div>
    <div style='grid-column:1/-1'><label>Remarks</label><textarea name='remarks'></textarea></div></div><br><button class='btn green'>Save Entry</button></form></div>""",members=members,today=date.today().isoformat())
    return page("Paid Return",body)

@app.route("/ledger")
def ledger():
    person=request.args.get("person","")
    c=db();rows=execute(c, "SELECT * FROM person_ledger WHERE person=? ORDER BY tx_date,id",(person,)).fetchall();c.close()
    paid=sum(r['paid'] or 0 for r in rows);ret=sum(r['returned'] or 0 for r in rows);bal=0;data=[]
    for r in rows: bal+=r['paid']-r['returned'];data.append((r,bal))
    body=render_template_string("""<div class='title'><div><h1>{{person}}</h1><small>Complete member ledger</small></div><div class='actions'><a class='btn green' href='/ledger/add'>+ Add Entry</a><a class='btn light' href='/people'>All Members</a></div></div>
    <div class='cards'><div class='card'><div class='label'>Total Paid</div><div class='num'>₹ {{'%.2f'|format(paid)}}</div></div><div class='card'><div class='label'>Total Return</div><div class='num'>₹ {{'%.2f'|format(ret)}}</div></div><div class='card'><div class='label'>Balance</div><div class='num'>₹ {{'%.2f'|format(paid-ret)}}</div></div></div>
    <div class='table'><table><tr><th>Date</th><th>Particular</th><th class='right'>Paid</th><th class='right'>Return</th><th class='right'>Balance</th><th>Remarks</th></tr>
    {%for r,b in data%}<tr><td>{{r.tx_date}}</td><td>{{r.particular}}</td><td class='right'>₹ {{'%.2f'|format(r.paid)}}</td><td class='right'>₹ {{'%.2f'|format(r.returned)}}</td><td class='right'><b>₹ {{'%.2f'|format(b)}}</b></td><td>{{r.remarks}}</td></tr>{%else%}<tr><td colspan='6'>No entries.</td></tr>{%endfor%}</table></div>""",person=person,paid=paid,ret=ret,data=data)
    return page("Ledger",body)

@app.route("/ledger/export")
def ledger_export():
    c=db();rows=execute(c, "SELECT * FROM person_ledger ORDER BY person,tx_date,id").fetchall();c.close()
    wb=Workbook();ws=wb.active;ws.title="Member Ledger";ws.append(["Member","Date","Particular","Paid","Return","Balance","Remarks"])
    balances={}
    for r in rows:
        balances[r['person']]=balances.get(r['person'],0)+r['paid']-r['returned']
        ws.append([r['person'],r['tx_date'],r['particular'],r['paid'],r['returned'],balances[r['person']],r['remarks']])
    path=os.path.join(BASE,"Member_Ledger.xlsx");wb.save(path)
    return send_file(path,as_attachment=True,download_name="Member_Ledger.xlsx")

@app.route("/manifest.webmanifest")
def manifest():
    return jsonify({"name":"Personal Expense Manager","short_name":"Expense Manager","start_url":"/","display":"standalone","background_color":"#f5f7fb","theme_color":"#111827","orientation":"portrait-primary","icons":[{"src":"/icon-192.png","sizes":"192x192","type":"image/png","purpose":"any maskable"},{"src":"/icon-512.png","sizes":"512x512","type":"image/png","purpose":"any maskable"}]})

@app.route("/sw.js")
def service_worker():
    js="const CACHE='expense-manager-v1';self.addEventListener('install',e=>{e.waitUntil(caches.open(CACHE).then(c=>c.addAll(['/','/manifest.webmanifest'])));self.skipWaiting()});self.addEventListener('activate',e=>{e.waitUntil(self.clients.claim())});self.addEventListener('fetch',e=>{if(e.request.method!=='GET')return;e.respondWith(fetch(e.request).catch(()=>caches.match(e.request)))});"
    return app.response_class(js,mimetype="application/javascript")

@app.route("/icon-192.png")
def icon192(): return send_file(os.path.join(BASE,"icon-192.png"),mimetype="image/png")
@app.route("/icon-512.png")
def icon512(): return send_file(os.path.join(BASE,"icon-512.png"),mimetype="image/png")

if __name__=="__main__":
    init()
    app.run(host="0.0.0.0",port=int(os.environ.get("PORT","5000")),debug=False)
