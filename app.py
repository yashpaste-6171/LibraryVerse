import os, io, sqlite3, secrets, smtplib
from datetime import date, timedelta
from email.message import EmailMessage
from functools import wraps
from flask import Flask, render_template, request, redirect, url_for, flash, session, g, send_file, jsonify
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
from flask_wtf.csrf import CSRFProtect
from PIL import Image
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from openpyxl import Workbook
import requests

ROOT = os.path.abspath(os.path.dirname(__file__))

DATA_DIR = os.environ.get("LIBRARY_DATA_DIR", ROOT)
os.makedirs(DATA_DIR, exist_ok=True)

DB = os.path.join(DATA_DIR, "library.db")

UPLOAD = os.path.join(DATA_DIR, "uploads")
os.makedirs(UPLOAD, exist_ok=True)
app=Flask(__name__)
def send_brevo_email(to_email, subject, html_content):
    api_key = os.environ.get("BREVO_API_KEY")
    sender_email = os.environ.get("MAIL_FROM")
    sender_name = os.environ.get("MAIL_FROM_NAME", "LibraryVerse")

    if not api_key or not sender_email or not to_email:
        app.logger.warning("Brevo email settings or recipient missing.")
        return False

    url = "https://api.brevo.com/v3/smtp/email"

    headers = {
        "accept": "application/json",
        "api-key": api_key,
        "content-type": "application/json"
    }

    payload = {
        "sender": {
            "name": sender_name,
            "email": sender_email
        },
        "to": [{"email": to_email}],
        "subject": subject,
        "htmlContent": html_content
    }

    try:
        response = requests.post(
            url,
            headers=headers,
            json=payload,
            timeout=15
        )
        response.raise_for_status()
        return True
    except Exception:
        app.logger.exception("Brevo email sending failed.")
        return False
def load_secret_key():
    configured=os.environ.get("SECRET_KEY")
    if configured: return configured
    key_file=os.path.join(ROOT,".secret_key")
    if os.path.exists(key_file):
        with open(key_file,"r",encoding="utf-8") as handle: return handle.read().strip()
    value=secrets.token_urlsafe(48)
    with open(key_file,"w",encoding="utf-8") as handle: handle.write(value)
    return value
app.config.update(SECRET_KEY=load_secret_key(),MAX_CONTENT_LENGTH=5*1024*1024,
                  SESSION_COOKIE_HTTPONLY=True,SESSION_COOKIE_SAMESITE="Lax")
CSRFProtect(app)
FINE_PER_DAY=5

def db():
    if "db" not in g:
        g.db=sqlite3.connect(DB); g.db.row_factory=sqlite3.Row; g.db.execute("PRAGMA foreign_keys=ON")
    return g.db

@app.teardown_appcontext
def close_db(_=None):
    c=g.pop("db",None)
    if c: c.close()


SAMPLE_BOOKS = [
    ("The Alchemist","Paulo Coelho","Fiction"),
    ("Atomic Habits","James Clear","Self-help"),
    ("Rich Dad Poor Dad","Robert T. Kiyosaki","Finance"),
    ("Wings of Fire","A. P. J. Abdul Kalam","Biography"),
    ("The Psychology of Money","Morgan Housel","Finance"),
    ("Harry Potter and the Philosopher's Stone","J. K. Rowling","Fantasy"),
    ("The Hobbit","J. R. R. Tolkien","Fantasy"),
    ("To Kill a Mockingbird","Harper Lee","Classic"),
    ("1984","George Orwell","Dystopian"),
    ("Pride and Prejudice","Jane Austen","Romance"),
    ("The Great Gatsby","F. Scott Fitzgerald","Classic"),
    ("The Little Prince","Antoine de Saint-Exupéry","Fiction"),
    ("Ikigai","Héctor García and Francesc Miralles","Self-help"),
    ("Deep Work","Cal Newport","Productivity"),
    ("Think and Grow Rich","Napoleon Hill","Finance"),
    ("The 7 Habits of Highly Effective People","Stephen R. Covey","Personal development"),
    ("How to Win Friends and Influence People","Dale Carnegie","Personal development"),
    ("Sapiens","Yuval Noah Harari","History"),
    ("A Brief History of Time","Stephen Hawking","Science"),
    ("The Theory of Everything","Stephen Hawking","Science"),
    ("The Diary of a Young Girl","Anne Frank","Biography"),
    ("The Blue Umbrella","Ruskin Bond","Fiction"),
    ("The Room on the Roof","Ruskin Bond","Fiction"),
    ("Malgudi Days","R. K. Narayan","Short stories"),
    ("The Guide","R. K. Narayan","Classic"),
    ("Train to Pakistan","Khushwant Singh","Historical fiction"),
    ("The God of Small Things","Arundhati Roy","Literary fiction"),
    ("The White Tiger","Aravind Adiga","Literary fiction"),
    ("Midnight's Children","Salman Rushdie","Historical fiction"),
    ("Gitanjali","Rabindranath Tagore","Poetry"),
    ("The Discovery of India","Jawaharlal Nehru","History"),
    ("India After Gandhi","Ramachandra Guha","History"),
    ("The Immortals of Meluha","Amish Tripathi","Fantasy"),
    ("The Secret of the Nagas","Amish Tripathi","Fantasy"),
    ("The Palace of Illusions","Chitra Banerjee Divakaruni","Mythology"),
    ("The Forest of Enchantments","Chitra Banerjee Divakaruni","Mythology"),
    ("The Silent Patient","Alex Michaelides","Thriller"),
    ("And Then There Were None","Agatha Christie","Mystery"),
    ("Murder on the Orient Express","Agatha Christie","Mystery"),
    ("The Da Vinci Code","Dan Brown","Thriller"),
    ("The Fault in Our Stars","John Green","Young adult"),
    ("Wonder","R. J. Palacio","Young adult"),
    ("The Hunger Games","Suzanne Collins","Dystopian"),
    ("The Catcher in the Rye","J. D. Salinger","Classic"),
    ("Jane Eyre","Charlotte Brontë","Classic"),
    ("Frankenstein","Mary Shelley","Gothic"),
    ("The Call of the Wild","Jack London","Adventure"),
    ("The Art of War","Sun Tzu","Strategy"),
    ("Man's Search for Meaning","Viktor E. Frankl","Psychology"),
    ("The Complete Poems","Emily Dickinson","Poetry"),
]


def make_cover(title, author, category, filename):
    """Create a local illustrated cover so every catalogue card has a working image."""
    from PIL import Image, ImageDraw, ImageFont
    palette = [(48,57,103),(91,73,145),(39,105,105),(148,86,71),(67,88,127),
               (113,83,117),(61,111,89),(133,101,54),(73,82,111),(115,74,74)]
    idx = sum(ord(ch) for ch in title) % len(palette)
    image = Image.new("RGB", (600, 820), palette[idx])
    draw = ImageDraw.Draw(image)
    font_path = os.path.join(os.environ.get("WINDIR", "C:/Windows"), "Fonts", "arial.ttf")
    try:
        title_font = ImageFont.truetype(font_path, 34)
        small_font = ImageFont.truetype(font_path, 17)
    except Exception:
        title_font = ImageFont.load_default()
        small_font = ImageFont.load_default()
    draw.rounded_rectangle((28,28,572,792), radius=24, outline=(230,225,255), width=3)
    draw.ellipse((390,65,660,335), outline=(220,215,255), width=5)
    draw.ellipse((-95,570,220,885), outline=(220,215,255), width=5)
    draw.text((62,82), "LIBRARYVERSE  /  READING COLLECTION", fill=(235,232,255), font=small_font)
    words, lines, line = title.split(), [], ""
    for word in words:
        candidate=(line+" "+word).strip()
        if draw.textbbox((0,0),candidate,font=title_font)[2] <= 455:
            line=candidate
        else:
            if line: lines.append(line)
            line=word
    if line: lines.append(line)
    y=285 if len(lines)<=3 else 230
    for part in lines[:6]:
        draw.text((62,y),part,fill="white",font=title_font); y+=53
    draw.line((62,y+18,275,y+18),fill=(225,220,255),width=4)
    draw.text((62,y+55),author[:44],fill=(235,232,255),font=small_font)
    draw.text((62,738),(category or "GENERAL").upper()[:28],fill=(235,232,255),font=small_font)
    image.save(os.path.join(UPLOAD,filename),"JPEG",quality=90)

def seed_sample_books(conn):
    """Seed 50 demo titles once and repair missing covers without replacing user records."""
    for i,(title,author,category) in enumerate(SAMPLE_BOOKS,start=1):
        isbn=f"DEMO-LIB-{i:04d}"
        row=conn.execute("SELECT id,cover_image FROM books WHERE isbn=?",(isbn,)).fetchone()
        if not row:
            filename=f"sample-cover-{i:02d}.jpg"
            make_cover(title,author,category,filename)
            total=2+(i%5)
            conn.execute("""INSERT INTO books(title,author,isbn,category,total_copies,available_copies,cover_image)
                            VALUES(?,?,?,?,?,?,?)""",(title,author,isbn,category,total,total,filename))
        else:
            filename=row["cover_image"] or f"sample-cover-{i:02d}.jpg"
            if not os.path.isfile(os.path.join(UPLOAD,filename)):
                filename=f"sample-cover-{i:02d}.jpg"
                make_cover(title,author,category,filename)
                conn.execute("UPDATE books SET cover_image=? WHERE id=?",(filename,row["id"]))
    # Repair cover references for pre-existing catalogue books too.
    rows=conn.execute("SELECT id,title,author,category,cover_image FROM books").fetchall()
    for row in rows:
        current=row["cover_image"]
        if not current or not os.path.isfile(os.path.join(UPLOAD,current)):
            filename=f"catalog-cover-{row['id']}.jpg"
            make_cover(row["title"],row["author"],row["category"] or "GENERAL",filename)
            conn.execute("UPDATE books SET cover_image=? WHERE id=?",(filename,row["id"]))

def init_db():
    c=sqlite3.connect(DB); c.row_factory=sqlite3.Row; c.execute("PRAGMA foreign_keys=ON")
    c.executescript("""
    CREATE TABLE IF NOT EXISTS librarians(id INTEGER PRIMARY KEY,username TEXT UNIQUE NOT NULL,password_hash TEXT NOT NULL,role TEXT NOT NULL DEFAULT 'librarian',active INTEGER DEFAULT 1);
    CREATE TABLE IF NOT EXISTS students(id INTEGER PRIMARY KEY,name TEXT NOT NULL,email TEXT UNIQUE,phone TEXT,student_code TEXT UNIQUE NOT NULL,password_hash TEXT,active INTEGER DEFAULT 1,created_at TEXT DEFAULT CURRENT_TIMESTAMP);
    CREATE TABLE IF NOT EXISTS books(id INTEGER PRIMARY KEY,title TEXT NOT NULL,author TEXT NOT NULL,isbn TEXT UNIQUE,category TEXT,total_copies INTEGER NOT NULL,available_copies INTEGER NOT NULL,cover_image TEXT,created_at TEXT DEFAULT CURRENT_TIMESTAMP);
    CREATE TABLE IF NOT EXISTS loans(id INTEGER PRIMARY KEY,book_id INTEGER NOT NULL REFERENCES books(id),student_id INTEGER NOT NULL REFERENCES students(id),issued_by INTEGER REFERENCES librarians(id),issue_date TEXT NOT NULL,due_date TEXT NOT NULL,return_date TEXT,fine REAL DEFAULT 0);
    CREATE TABLE IF NOT EXISTS reservations(id INTEGER PRIMARY KEY,book_id INTEGER NOT NULL REFERENCES books(id),student_id INTEGER NOT NULL REFERENCES students(id),reserved_at TEXT NOT NULL,status TEXT DEFAULT 'Waiting',notified_at TEXT,fulfilled_at TEXT);
    """)
    # Migrate databases created by earlier LibraryDesk versions.
    for table,col,definition in [("students","password_hash","TEXT"),("students","active","INTEGER DEFAULT 1"),("books","cover_image","TEXT"),("loans","issued_by","INTEGER REFERENCES librarians(id)")]:
        cols={r[1] for r in c.execute(f"PRAGMA table_info({table})")}
        if col not in cols: c.execute(f"ALTER TABLE {table} ADD COLUMN {col} {definition}")
    # Carry forward a legacy admin hash if a v1 admins table exists.
    if c.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='admins'").fetchone():
        old=c.execute("SELECT username,password_hash FROM admins LIMIT 1").fetchone()
        if old and not c.execute("SELECT 1 FROM librarians WHERE username=?",(old[0],)).fetchone():
            c.execute("INSERT INTO librarians(username,password_hash,role) VALUES(?,?,?)",(old[0],old[1],"admin"))
    if not c.execute("SELECT 1 FROM librarians LIMIT 1").fetchone():
        c.execute("INSERT INTO librarians(username,password_hash,role) VALUES(?,?,?)",
          (os.environ.get("LIBRARY_ADMIN_USER","admin"),generate_password_hash(os.environ.get("LIBRARY_ADMIN_PASSWORD","ChangeMe123!")),"admin"))
    seed_sample_books(c)
    c.commit(); c.close()

def librarian_required(fn):
    @wraps(fn)
    def wrap(*a,**kw):
        if not session.get("librarian_id"): return redirect(url_for("login"))
        return fn(*a,**kw)
    return wrap

def admin_required(fn):
    @wraps(fn)
    @librarian_required
    def wrap2(*a,**kw):
        if session.get("role")!="admin": flash("Administrator access required.","danger"); return redirect(url_for("admin_dashboard"))
        return fn(*a,**kw)
    return wrap2

def student_required(fn):
    @wraps(fn)
    def wrap(*a, **kw):
        sid = session.get("student_id")
        if not sid:
            flash("Please sign in with your student account.", "info")
            return redirect(url_for("student_login"))

        # Validate the session against the active database. Old browser sessions
        # can contain an ID from a previous database copy.
        c = db()
        student = c.execute(
            "SELECT id, name, student_code, active FROM students WHERE id=?", (sid,)
        ).fetchone()
        if student is None and session.get("student_code"):
            student = c.execute(
                "SELECT id, name, student_code, active FROM students WHERE student_code=?",
                (session.get("student_code"),)
            ).fetchone()
        if student is None or not student["active"]:
            session.clear()
            flash("Your student session could not be matched to an active library account. Please sign in again.", "warning")
            return redirect(url_for("student_login"))

        # Refresh the session values from the canonical database row.
        session["student_id"] = student["id"]
        session["student_name"] = student["name"]
        session["student_code"] = student["student_code"]
        return fn(*a, **kw)
    return wrap

def send_mail(to,subject,body):
    host=os.environ.get("SMTP_HOST")
    if not host: raise RuntimeError("SMTP is not configured. Set SMTP_HOST, SMTP_PORT, SMTP_USER, SMTP_PASSWORD.")
    m=EmailMessage(); m["Subject"]=subject; m["From"]=os.environ.get("MAIL_FROM",os.environ.get("SMTP_USER")); m["To"]=to; m.set_content(body)
    with smtplib.SMTP(host,int(os.environ.get("SMTP_PORT","587")),timeout=20) as s:
        s.starttls(); s.login(os.environ["SMTP_USER"],os.environ["SMTP_PASSWORD"]); s.send_message(m)

def save_cover(f):
    if not f or not f.filename: return None
    ext=f.filename.rsplit(".",1)[-1].lower()
    if ext not in {"png","jpg","jpeg","webp"}: raise ValueError("Upload PNG, JPG or WEBP cover images only.")
    im=Image.open(f.stream).convert("RGB"); im.thumbnail((700,900))
    name=secrets.token_hex(12)+".jpg"; im.save(os.path.join(UPLOAD,name),"JPEG",quality=88); return name

@app.template_filter("money")
def money(v): return f"{float(v or 0):.2f}"

@app.context_processor
def globals_():
    return {"today":date.today().isoformat(),"fine_per_day":FINE_PER_DAY}

@app.route("/")
def home():
    if session.get("librarian_id"): return redirect(url_for("admin_dashboard"))
    if session.get("student_id"): return redirect(url_for("student_dashboard"))
    return redirect(url_for("login"))

@app.route("/login",methods=["GET","POST"])
def login():
    if request.method=="POST":
        u=db().execute("SELECT * FROM librarians WHERE username=? AND active=1",(request.form.get("username","").strip(),)).fetchone()
        if u and check_password_hash(u["password_hash"],request.form.get("password","")):
            session.clear(); session.update(librarian_id=u["id"],username=u["username"],role=u["role"])
            return redirect(url_for("admin_dashboard"))
        flash("Incorrect administrator/librarian username or password.","danger")
    return render_template("login.html")

@app.route("/student/login",methods=["GET","POST"])
def student_login():
    if request.method=="POST":
        code=request.form.get("student_code","").strip()
        s=db().execute("SELECT * FROM students WHERE student_code=? AND active=1",(code,)).fetchone()
        if s and s["password_hash"] and check_password_hash(s["password_hash"],request.form.get("password","")):
            session.clear(); session.update(student_id=s["id"],student_name=s["name"],student_code=s["student_code"],role="student")
            return redirect(url_for("student_dashboard"))
        flash("Student ID or password is incorrect. If this is your first visit, create a student account.","danger")
    return render_template("student_login.html")

@app.route("/student/register",methods=["GET","POST"])
def student_register():
    if request.method=="POST":
        name=request.form.get("name","").strip(); code=request.form.get("student_code","").strip()
        email=request.form.get("email","").strip().lower(); phone=request.form.get("phone","").strip()
        pw=request.form.get("password",""); confirm=request.form.get("confirm_password","")
        if not name or not code or not email: flash("Name, student ID and email are required.","danger")
        elif len(pw)<10: flash("Password must contain at least 10 characters.","danger")
        elif pw!=confirm: flash("Passwords do not match.","danger")
        else:
            c=db(); existing=c.execute("SELECT * FROM students WHERE student_code=?",(code,)).fetchone()
            try:
                if existing:
                    if existing["password_hash"]: raise ValueError("An account already exists for this student ID. Please sign in.")
                    if (existing["email"] or "").lower()!=email: raise ValueError("Student ID found, but email does not match the library record. Contact the librarian.")
                    c.execute("UPDATE students SET name=?,phone=?,password_hash=? WHERE id=?",
                              (name,phone,generate_password_hash(pw),existing["id"]))
                    sid=existing["id"]
                else:
                    cur=c.execute("INSERT INTO students(name,email,phone,student_code,password_hash) VALUES(?,?,?,?,?)",
                                  (name,email,phone,code,generate_password_hash(pw))); sid=cur.lastrowid
                c.commit()
                 send_brevo_email(
                    email,
                    "Welcome to LibraryVerse!",
                    f"""
                    <h2>Welcome to LibraryVerse, {name}!</h2>
                    <p>Your student account has been created successfully.</p>
                    <p><b>Student ID:</b> {code}</p>
                    <p>You can now sign in and explore your library.</p>
                    <p>Happy reading!<br>LibraryVerse Team</p>
                    """
                )


                session.clear()
                session.update(
                    student_id=sid,
                    student_name=name,
                    student_code=code,
                    role="student"
                )
                return redirect(url_for("student_dashboard"))    
            except (sqlite3.IntegrityError,ValueError) as e:
                flash(str(e) or "Email or student ID is already registered.","danger")
    return render_template("student_register.html")

@app.route("/logout")
def logout():
    session.clear(); flash("You have signed out.","success"); return redirect(url_for("home"))

@app.route("/admin")
@librarian_required
def admin_dashboard():
    c=db()
    stats={"titles":c.execute("SELECT COUNT(*) FROM books").fetchone()[0],
           "copies":c.execute("SELECT COALESCE(SUM(total_copies),0) FROM books").fetchone()[0],
           "available":c.execute("SELECT COALESCE(SUM(available_copies),0) FROM books").fetchone()[0],
           "students":c.execute("SELECT COUNT(*) FROM students").fetchone()[0],
           "issued":c.execute("SELECT COUNT(*) FROM loans WHERE return_date IS NULL").fetchone()[0],
           "overdue":c.execute("SELECT COUNT(*) FROM loans WHERE return_date IS NULL AND due_date<?",(date.today().isoformat(),)).fetchone()[0],
           "reservations":c.execute("SELECT COUNT(*) FROM reservations WHERE status='Waiting'").fetchone()[0]}
    recent=c.execute("""SELECT l.*,b.title,s.name student_name FROM loans l JOIN books b ON b.id=l.book_id JOIN students s ON s.id=l.student_id ORDER BY l.id DESC LIMIT 8""").fetchall()
    monthly=c.execute("""SELECT substr(issue_date,1,7) month,COUNT(*) total FROM loans
                         GROUP BY substr(issue_date,1,7) ORDER BY month DESC LIMIT 6""").fetchall()
    monthly=list(reversed(monthly))
    categories=c.execute("""SELECT COALESCE(NULLIF(category,''),'General') label,COUNT(*) total
                            FROM books GROUP BY COALESCE(NULLIF(category,''),'General')
                            ORDER BY total DESC LIMIT 6""").fetchall()
    return render_template("admin_dashboard.html",stats=stats,recent=recent,monthly=monthly,categories=categories)

@app.route("/student")
@student_required
def student_dashboard():
    c=db(); sid=session["student_id"]
    s=c.execute("SELECT * FROM students WHERE id=? AND active=1",(sid,)).fetchone()
    if s is None:
        session.clear()
        flash("Student account was not found. Please sign in again.", "warning")
        return redirect(url_for("student_login"))
    active=c.execute("""SELECT l.*,b.title,b.author,b.cover_image FROM loans l JOIN books b ON b.id=l.book_id WHERE l.student_id=? AND l.return_date IS NULL ORDER BY l.due_date""",(sid,)).fetchall()
    history=c.execute("""SELECT l.*,b.title,b.author,b.cover_image FROM loans l JOIN books b ON b.id=l.book_id WHERE l.student_id=? ORDER BY l.id DESC LIMIT 12""",(sid,)).fetchall()
    reservations=c.execute("""SELECT r.*,b.title FROM reservations r JOIN books b ON b.id=r.book_id WHERE r.student_id=? ORDER BY r.id DESC LIMIT 5""",(sid,)).fetchall()
    returned=c.execute("SELECT COUNT(*) FROM loans WHERE student_id=? AND return_date IS NOT NULL",(sid,)).fetchone()[0]
    overdue=sum(1 for loan in active if loan["due_date"] < date.today().isoformat())
    monthly=c.execute("""SELECT substr(issue_date,1,7) month,COUNT(*) total FROM loans
                         WHERE student_id=? GROUP BY substr(issue_date,1,7) ORDER BY month DESC LIMIT 6""",(sid,)).fetchall()
    monthly=list(reversed(monthly))
    return render_template("student_dashboard.html",student=s,active=active,history=history,reservations=reservations,
                           returned=returned,overdue=overdue,monthly=monthly)

@app.route("/student/catalog")
@student_required
def student_catalog():
    q=request.args.get("q","").strip()
    books=db().execute("""SELECT * FROM books WHERE title LIKE ? OR author LIKE ? OR COALESCE(category,'') LIKE ? ORDER BY title""",
      (f"%{q}%",f"%{q}%",f"%{q}%")).fetchall()
    return render_template("student_catalog.html",books=books,q=q)

@app.route("/student/borrow/<int:book_id>",methods=["POST"])
@student_required
def student_borrow(book_id):
    """Let a signed-in student borrow an available copy from the catalogue."""
    c = db()
    try:
        c.execute("BEGIN IMMEDIATE")
        book = c.execute("SELECT * FROM books WHERE id=?", (book_id,)).fetchone()
        if not book:
            raise ValueError("Book not found.")
        if book["available_copies"] < 1:
            raise ValueError("This book is no longer available. Please reserve it instead.")
        already = c.execute(
            "SELECT 1 FROM loans WHERE student_id=? AND book_id=? AND return_date IS NULL",
            (session["student_id"], book_id)
        ).fetchone()
        if already:
            raise ValueError("You already have this book on loan.")
        issue = date.today()
        due = issue + timedelta(days=14)
        c.execute(
            "INSERT INTO loans(book_id,student_id,issued_by,issue_date,due_date) VALUES(?,?,?,?,?)",
            (book_id, session["student_id"], None, issue.isoformat(), due.isoformat())
        )
        c.execute("UPDATE books SET available_copies=available_copies-1 WHERE id=?", (book_id,))
        c.commit()
        flash(f"'{book['title']}' has been borrowed. Please return it by {due.isoformat()}.", "success")
        return redirect(url_for("student_dashboard"))
    except ValueError as e:
        c.rollback()
        flash(str(e), "warning")
        return redirect(url_for("student_catalog"))
    except Exception:
        c.rollback()
        raise


@app.route("/student/return/<int:loan_id>", methods=["POST"])
@student_required
def student_return(loan_id):
    c=db()
    try:
        c.execute("BEGIN IMMEDIATE")
        loan=c.execute("""SELECT l.*,b.title FROM loans l JOIN books b ON b.id=l.book_id
                          WHERE l.id=? AND l.student_id=? AND l.return_date IS NULL""",
                       (loan_id,session["student_id"])).fetchone()
        if not loan:
            raise ValueError("This loan is not active or does not belong to your account.")
        returned=date.today()
        fine=max(0,(returned-date.fromisoformat(loan["due_date"])).days)*FINE_PER_DAY
        updated=c.execute("UPDATE loans SET return_date=?,fine=? WHERE id=? AND return_date IS NULL",
                          (returned.isoformat(),fine,loan_id))
        if updated.rowcount != 1:
            raise ValueError("This book has already been returned.")
        c.execute("UPDATE books SET available_copies=MIN(total_copies,available_copies+1) WHERE id=?",
                  (loan["book_id"],))
        c.commit()
        if fine:
            flash(f"'{loan['title']}' was returned. An overdue fine of ₹{fine:.2f} has been recorded.", "warning")
        else:
            flash(f"'{loan['title']}' was returned successfully. Thank you!", "success")
    except ValueError as e:
        c.rollback()
        flash(str(e),"danger")
    except Exception:
        c.rollback()
        raise
    return redirect(url_for("student_dashboard"))

@app.route("/student/reserve/<int:book_id>",methods=["POST"])
@student_required
def student_reserve(book_id):
    c=db(); book=c.execute("SELECT * FROM books WHERE id=?",(book_id,)).fetchone()
    if not book: flash("Book not found.","danger")
    elif book["available_copies"]>0: flash("This book is currently available. Please ask the librarian to issue it.","info")
    elif c.execute("SELECT 1 FROM reservations WHERE book_id=? AND student_id=? AND status IN ('Waiting','Ready')",(book_id,session["student_id"])).fetchone():
        flash("You already have an active reservation for this book.","warning")
    else:
        c.execute("INSERT INTO reservations(book_id,student_id,reserved_at) VALUES(?,?,?)",(book_id,session["student_id"],date.today().isoformat()))
        c.commit(); flash("Reservation placed successfully.","success")
    return redirect(url_for("student_catalog"))

@app.route("/student/profile",methods=["GET","POST"])
@student_required
def student_profile():
    c=db(); s=c.execute("SELECT * FROM students WHERE id=?",(session["student_id"],)).fetchone()
    if request.method=="POST":
        name=request.form.get("name","").strip(); phone=request.form.get("phone","").strip()
        new_password=request.form.get("new_password","")
        if new_password and len(new_password)<10:
            flash("New password must contain at least 10 characters.","danger")
        elif name:
            c.execute("UPDATE students SET name=?,phone=? WHERE id=?",(name,phone,s["id"]))
            if new_password:
                c.execute("UPDATE students SET password_hash=? WHERE id=?",(generate_password_hash(new_password),s["id"]))
            c.commit()
            session["student_name"]=name; flash("Profile updated.","success"); return redirect(url_for("student_profile"))
    import base64, qrcode
    qr=qrcode.make(f"LIBRARYVERSE|STUDENT|{s['student_code']}")
    buffer=io.BytesIO(); qr.save(buffer,format="PNG")
    qr_data="data:image/png;base64,"+base64.b64encode(buffer.getvalue()).decode("ascii")
    return render_template("student_profile.html",student=s,qr_data=qr_data)

@app.route("/books",methods=["GET","POST"])
@librarian_required
def books():
    c=db()
    if request.method=="POST":
        try:
            title=request.form.get("title","").strip(); author=request.form.get("author","").strip()
            n=int(request.form.get("total_copies","1")); isbn=request.form.get("isbn","").strip() or None
            cover=save_cover(request.files.get("cover"))
            if not title or not author or n<1: raise ValueError("Enter title, author and at least one copy.")
            c.execute("INSERT INTO books(title,author,isbn,category,total_copies,available_copies,cover_image) VALUES(?,?,?,?,?,?,?)",
                      (title,author,isbn,request.form.get("category","").strip(),n,n,cover)); c.commit(); flash("Book added.","success")
        except ValueError as e: flash(str(e),"danger")
        except sqlite3.IntegrityError: flash("ISBN already exists.","danger")
        return redirect(url_for("books"))
    q=request.args.get("q","").strip()
    rows=c.execute("SELECT * FROM books WHERE title LIKE ? OR author LIKE ? OR COALESCE(isbn,'') LIKE ? ORDER BY title",
                   (f"%{q}%",f"%{q}%",f"%{q}%")).fetchall()
    return render_template("books.html",books=rows,q=q)


@app.route("/books/<int:bid>/edit", methods=["GET","POST"])
@librarian_required
def edit_book(bid):
    c=db()
    book=c.execute("SELECT * FROM books WHERE id=?",(bid,)).fetchone()
    if not book:
        flash("Book not found.","danger")
        return redirect(url_for("books"))
    if request.method=="POST":
        try:
            title=request.form.get("title","").strip()
            author=request.form.get("author","").strip()
            total=int(request.form.get("total_copies","1"))
            isbn=request.form.get("isbn","").strip() or None
            if not title or not author or total<1: raise ValueError("Enter a title, author and at least one copy.")
            issued=c.execute("SELECT COUNT(*) FROM loans WHERE book_id=? AND return_date IS NULL",(bid,)).fetchone()[0]
            if total<issued: raise ValueError(f"Total copies cannot be less than {issued} copies currently on loan.")
            cover=book["cover_image"]
            uploaded=request.files.get("cover")
            if uploaded and uploaded.filename: cover=save_cover(uploaded)
            c.execute("""UPDATE books SET title=?,author=?,isbn=?,category=?,total_copies=?,
                         available_copies=?,cover_image=? WHERE id=?""",
                      (title,author,isbn,request.form.get("category","").strip(),total,total-issued,cover,bid))
            c.commit(); flash("Book details updated.","success")
            return redirect(url_for("books"))
        except ValueError as e: flash(str(e),"danger")
        except sqlite3.IntegrityError: flash("That ISBN is already assigned to another book.","danger")
    return render_template("edit_book.html",book=book)

@app.route("/books/<int:bid>/delete",methods=["POST"])
@librarian_required
def delete_book(bid):
    c=db()
    used=c.execute("SELECT (SELECT COUNT(*) FROM loans WHERE book_id=?)+(SELECT COUNT(*) FROM reservations WHERE book_id=?)",(bid,bid)).fetchone()[0]
    if used: flash("Cannot delete a book with loan or reservation history.","danger")
    else: c.execute("DELETE FROM books WHERE id=?",(bid,)); c.commit(); flash("Book deleted.","success")
    return redirect(url_for("books"))

@app.route("/students",methods=["GET","POST"])
@librarian_required
def students():
    c=db()
    if request.method=="POST":
        try:
            c.execute("INSERT INTO students(name,email,phone,student_code) VALUES(?,?,?,?)",
             (request.form["name"].strip(),request.form.get("email","").strip() or None,request.form.get("phone","").strip(),request.form["student_code"].strip()))
            c.commit(); flash("Student added. They can now create an account using the same student ID and registered email.","success")
        except sqlite3.IntegrityError: flash("Student ID or email already exists.","danger")
        return redirect(url_for("students"))
    q=request.args.get("q","").strip()
    rows=c.execute("SELECT * FROM students WHERE name LIKE ? OR student_code LIKE ? ORDER BY name",(f"%{q}%",f"%{q}%")).fetchall()
    return render_template("students.html",students=rows,q=q)

@app.route("/loans",methods=["GET","POST"])
@librarian_required
def loans():
    c=db()
    if request.method=="POST":
        try:
            bid=int(request.form["book_id"]); sid=int(request.form["student_id"])
            issue=request.form.get("issue_date") or date.today().isoformat(); due=request.form["due_date"]
            if due<issue: raise ValueError()
            c.execute("BEGIN IMMEDIATE")
            b=c.execute("SELECT available_copies FROM books WHERE id=?",(bid,)).fetchone()
            if not b or b["available_copies"]<1: raise ValueError()
            c.execute("INSERT INTO loans(book_id,student_id,issued_by,issue_date,due_date) VALUES(?,?,?,?,?)",
                      (bid,sid,session["librarian_id"],issue,due))
            c.execute("UPDATE books SET available_copies=available_copies-1 WHERE id=?",(bid,)); c.commit(); flash("Book issued.","success")
        except (ValueError,KeyError): c.rollback(); flash("Select an available book, student and valid dates.","danger")
        return redirect(url_for("loans"))
    rows=c.execute("""SELECT l.*,b.title,s.name student_name,s.student_code FROM loans l JOIN books b ON b.id=l.book_id JOIN students s ON s.id=l.student_id ORDER BY l.id DESC""").fetchall()
    return render_template("loans.html",loans=rows,books=c.execute("SELECT * FROM books WHERE available_copies>0").fetchall(),
                           students=c.execute("SELECT * FROM students ORDER BY name").fetchall())

@app.route("/loans/<int:lid>/return",methods=["POST"])
@librarian_required
def return_loan(lid):
    c=db(); l=c.execute("SELECT * FROM loans WHERE id=? AND return_date IS NULL",(lid,)).fetchone()
    if not l: flash("Loan not found or already returned.","danger")
    else:
        fine=max(0,(date.today()-date.fromisoformat(l["due_date"])).days)*FINE_PER_DAY
        c.execute("UPDATE loans SET return_date=?,fine=? WHERE id=?",(date.today().isoformat(),fine,lid))
        c.execute("UPDATE books SET available_copies=available_copies+1 WHERE id=?",(l["book_id"],)); c.commit()
        flash(f"Book returned. Fine: ₹{fine:.2f}","success")
    return redirect(url_for("loans"))

@app.route("/reservations")
@librarian_required
def reservations():
    rows=db().execute("""SELECT r.*,b.title,s.name student_name,s.student_code FROM reservations r JOIN books b ON b.id=r.book_id JOIN students s ON s.id=r.student_id ORDER BY r.id DESC""").fetchall()
    return render_template("reservations.html",reservations=rows)

@app.route("/reservations/<int:rid>/ready",methods=["POST"])
@librarian_required
def reservation_ready(rid):
    c=db(); r=c.execute("""SELECT r.*,s.email,s.name,b.title FROM reservations r JOIN students s ON s.id=r.student_id JOIN books b ON b.id=r.book_id WHERE r.id=?""",(rid,)).fetchone()
    if r:
        c.execute("UPDATE reservations SET status='Ready' WHERE id=?",(rid,)); c.commit()
        if r["email"]:
            try:
                send_mail(r["email"],"Your library reservation is ready",f"Hello {r['name']}, your reserved book '{r['title']}' is ready for collection.")
                flash("Marked ready and email sent.","success")
            except Exception as e: flash(f"Marked ready, but email was not sent: {e}","warning")
        else: flash("Marked ready. Student has no email address.","success")
    return redirect(url_for("reservations"))

@app.route("/reports")
@librarian_required
def reports():
    c=db(); rows=c.execute("""SELECT l.*,b.title,s.name student_name FROM loans l JOIN books b ON b.id=l.book_id JOIN students s ON s.id=l.student_id ORDER BY l.id DESC""").fetchall()
    return render_template("reports.html",loans=rows,total_fines=c.execute("SELECT COALESCE(SUM(fine),0) FROM loans").fetchone()[0])

@app.route("/reports/excel")
@librarian_required
def excel():
    rows=db().execute("""SELECT l.id,b.title,s.name,s.student_code,l.issue_date,l.due_date,l.return_date,l.fine FROM loans l JOIN books b ON b.id=l.book_id JOIN students s ON s.id=l.student_id""").fetchall()
    w=Workbook(); sh=w.active; sh.title="Loan report"; sh.append(["Loan ID","Book","Student","Student ID","Issued","Due","Returned","Fine"])
    for r in rows: sh.append(list(r))
    f=io.BytesIO(); w.save(f); f.seek(0)
    return send_file(f,as_attachment=True,download_name="library-report.xlsx")
@app.route("/reports/pdf")
@librarian_required
def pdf():
    rows=db().execute("""SELECT l.id,b.title,s.name,l.issue_date,l.due_date,l.return_date,l.fine FROM loans l JOIN books b ON b.id=l.book_id JOIN students s ON s.id=l.student_id""").fetchall()
    f=io.BytesIO(); c=canvas.Canvas(f,pagesize=A4); y=800
    c.setFont("Helvetica-Bold",16); c.drawString(40,y,"Library Loan Report"); y-=35
    c.setFont("Helvetica",8)
    for r in rows:
        c.drawString(40,y,f"#{r['id']}  {r['title'][:28]}  {r['name'][:18]}  {r['issue_date']}  Due {r['due_date']}  Fine {r['fine']:.2f}")
        y-=15
        if y<45: c.showPage(); y=800; c.setFont("Helvetica",8)
    c.save(); f.seek(0); return send_file(f,as_attachment=True,download_name="library-report.pdf")

@app.route("/scan")
@librarian_required
def scan(): return render_template("scan.html")
@app.route("/scan/lookup")
@librarian_required
def scan_lookup():
    code=request.args.get("code","").strip(); c=db()
    b=c.execute("SELECT * FROM books WHERE isbn=?",(code,)).fetchone()
    s=c.execute("SELECT * FROM students WHERE student_code=?",(code,)).fetchone()
    if b: return jsonify(type="Book",name=b["title"],available=b["available_copies"])
    if s: return jsonify(type="Student",name=s["name"],student_id=s["student_code"])
    return jsonify(error="No matching ISBN or student ID."),404

@app.route("/change-password",methods=["GET","POST"])
@librarian_required
def change_password():
    c=db(); u=c.execute("SELECT * FROM librarians WHERE id=?",(session["librarian_id"],)).fetchone()
    if request.method=="POST":
        old=request.form.get("current",""); new=request.form.get("new","")
        if not check_password_hash(u["password_hash"],old): flash("Current password is incorrect.","danger")
        elif len(new)<10: flash("Use at least 10 characters.","danger")
        else:
            c.execute("UPDATE librarians SET password_hash=? WHERE id=?",(generate_password_hash(new),u["id"])); c.commit(); flash("Password changed.","success")
    return render_template("change_password.html")

@app.route("/librarians",methods=["GET","POST"])
@admin_required
def librarians():
    c=db()
    if request.method=="POST":
        u=request.form.get("username","").strip(); p=request.form.get("password",""); role=request.form.get("role","librarian")
        if len(u)<3 or len(p)<10: flash("Username 3+ chars; password 10+ chars.","danger")
        else:
            try: c.execute("INSERT INTO librarians(username,password_hash,role) VALUES(?,?,?)",(u,generate_password_hash(p),role)); c.commit(); flash("Account created.","success")
            except sqlite3.IntegrityError: flash("Username already exists.","danger")
        return redirect(url_for("librarians"))
    return render_template("librarians.html",users=c.execute("SELECT * FROM librarians").fetchall())

@app.route("/reminders",methods=["POST"])
@admin_required
def reminders():
    rows=db().execute("""SELECT l.due_date,s.name,s.email,b.title FROM loans l JOIN students s ON s.id=l.student_id JOIN books b ON b.id=l.book_id WHERE l.return_date IS NULL AND s.email IS NOT NULL AND l.due_date<=?""",((date.today()+timedelta(days=2)).isoformat(),)).fetchall()
    sent=0
    for r in rows:
        try: send_mail(r["email"],"Library due-date reminder",f"Hello {r['name']}, '{r['title']}' is due on {r['due_date']}. Please return it to the library."); sent+=1
        except Exception: pass
    flash(f"Sent {sent} reminder email(s). Check SMTP settings if none were sent.","info"); return redirect(url_for("reports"))

if __name__=="__main__":
    init_db()
    app.run(debug=os.environ.get("FLASK_DEBUG","0")=="1")
