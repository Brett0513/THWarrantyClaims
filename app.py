from flask import Flask, render_template, request, redirect, url_for, flash, send_from_directory, jsonify
from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager, login_user, login_required, logout_user, current_user, UserMixin
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
from datetime import datetime, timedelta
import calendar
import sys
from ics import Calendar, Event
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
import pytz
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

app = Flask(__name__)
app.config['SECRET_KEY'] = 'your_super_secret_key'
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///warranty.db'
app.config['UPLOAD_FOLDER'] = os.path.join('static', 'uploads')
app.config['MAX_CONTENT_LENGTH'] = 10 * 1024 * 1024
app.config['ICS_FOLDER'] = os.path.join('static', 'ics')

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['ICS_FOLDER'], exist_ok=True)

db = SQLAlchemy(app)
login_manager = LoginManager(app)
login_manager.login_view = 'login'

from datetime import datetime
import pytz

def cst_now():
    """Returns the current datetime in Central Time."""
    return datetime.now(pytz.timezone('America/Chicago'))

class User(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    email = db.Column(db.String(100), unique=True, nullable=False)
    name = db.Column(db.String(100), nullable=False)
    password = db.Column(db.String(200), nullable=False)

class Vendor(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False)
    contact_number = db.Column(db.String(50))
    email = db.Column(db.String(100))

class Assignee(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False)
    contact_number = db.Column(db.String(50))
    email = db.Column(db.String(100))

class Claim(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    address = db.Column(db.String(200), nullable=False)
    homeowner_name = db.Column(db.String(100))
    homeowner_email = db.Column(db.String(100))
    homeowner_phone = db.Column(db.String(50))
    cobuyer_name = db.Column(db.String(100))
    cobuyer_email = db.Column(db.String(100))
    cobuyer_phone = db.Column(db.String(50))
    warranty_type = db.Column(db.String(100))
    issue_description = db.Column(db.Text)
    date_reported = db.Column(db.Date, default=cst_now)
    status = db.Column(db.String(20), default='Open')
    assignee_id = db.Column(db.Integer, db.ForeignKey('assignee.id'))
    assignee = db.relationship('Assignee')
    closures = db.relationship('ClaimClosure', backref='claim', cascade='all, delete-orphan')
    logs = db.relationship('ClaimLog', backref='claim', cascade='all, delete-orphan')
    workorders = db.relationship('WorkOrder', backref='claim', cascade='all, delete-orphan')

class ClaimPhoto(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    claim_id = db.Column(db.Integer, db.ForeignKey('claim.id'), nullable=False)
    filename = db.Column(db.String(200), nullable=False)
    claim = db.relationship('Claim', backref='photos')

class WorkOrder(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    claim_id = db.Column(db.Integer, db.ForeignKey('claim.id'), nullable=False)
    vendor_id = db.Column(db.Integer, db.ForeignKey('vendor.id'))
    assignee_id = db.Column(db.Integer, db.ForeignKey('assignee.id'))
    scheduled_date = db.Column(db.Date)
    scheduled_time = db.Column(db.String(8))
    status = db.Column(db.String(20), default='Scheduled')
    notes = db.Column(db.String(200))
    vendor = db.relationship('Vendor')
    assignee = db.relationship('Assignee')

class ClaimLog(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    claim_id = db.Column(db.Integer, db.ForeignKey('claim.id'), nullable=False)
    user_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=False)
    timestamp = db.Column(db.DateTime, default=cst_now)
    action = db.Column(db.String(200), nullable=False)
    user = db.relationship('User')

class ClaimClosure(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    claim_id = db.Column(db.Integer, db.ForeignKey('claim.id'), nullable=False)
    reasons = db.Column(db.String(400))
    notes = db.Column(db.Text)
    timestamp = db.Column(db.DateTime, default=cst_now)


@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))

# Utility function for PDF generation
def generate_workorder_pdf(claim, workorder, vendor=None, assignee=None, pdf_path="workorder.pdf"):
    c = canvas.Canvas(pdf_path, pagesize=letter)
    c.setFont("Helvetica-Bold", 16)
    c.drawString(72, 720, "Work Order")
    c.setFont("Helvetica", 12)
    y = 700
    lines = [
        f"Claim ID: {claim.id}",
        f"Address: {claim.address}",
        f"Homeowner: {claim.homeowner_name or ''}",
        f"Warranty Type: {claim.warranty_type or ''}",
        f"Issue: {claim.issue_description or ''}",
        f"Date Reported: {claim.date_reported.strftime('%Y-%m-%d') if claim.date_reported else ''}",
        "",
        f"Assigned Vendor: {vendor.name if vendor else 'N/A'}",
        f"Vendor Contact: {vendor.contact_number if vendor else 'N/A'}",
        f"Vendor Email: {vendor.email if vendor else 'N/A'}",
        "",
        f"Assigned Assignee: {assignee.name if assignee else 'N/A'}",
        f"Assignee Contact: {assignee.contact_number if assignee else 'N/A'}",
        f"Assignee Email: {assignee.email if assignee else 'N/A'}",
        "",
        f"Scheduled Date: {workorder.scheduled_date.strftime('%Y-%m-%d') if workorder.scheduled_date else ''}",
        f"Scheduled Time: {workorder.scheduled_time or ''}",
        f"Notes: {workorder.notes or ''}",
        f"Status: {workorder.status or ''}",
    ]
    for line in lines:
        c.drawString(72, y, line)
        y -= 20
    c.save()
    return pdf_path

@app.route('/')
@login_required
def index():
    claims = Claim.query.all()
    workorders = WorkOrder.query.all()

    claim_assignments = {}
    claim_status_override = {}
    claim_scheduled_dates = {}
    now = datetime.now().date()
    for claim in claims:
        wos = [wo for wo in workorders if wo.claim_id == claim.id]
        # NEW LOGIC STARTS HERE
        if claim.status == "Deferred":
            claim_status_override[claim.id] = "Deferred"
        elif claim.status == "Closed":
            claim_status_override[claim.id] = "Closed"
        else:
            is_scheduled = any(wo.scheduled_date and wo.scheduled_date >= now for wo in wos)
            if is_scheduled:
                claim_status_override[claim.id] = "Scheduled"
            else:
                claim_status_override[claim.id] = "Open"
        # NEW LOGIC ENDS HERE

        # Scheduled date/assignment
        if wos:
            scheduled_wos = [wo for wo in wos if wo.scheduled_date]
            if scheduled_wos:
                latest_sched = max(scheduled_wos, key=lambda wo: wo.scheduled_date)
                claim_scheduled_dates[claim.id] = latest_sched.scheduled_date.strftime('%Y-%m-%d')
                assigned = []
                if latest_sched.assignee:
                    assigned.append(latest_sched.assignee.name)
                if latest_sched.vendor:
                    assigned.append(latest_sched.vendor.name)
                claim_assignments[claim.id] = ', '.join(assigned) if assigned else "Unassigned"
            else:
                claim_scheduled_dates[claim.id] = ""
                claim_assignments[claim.id] = "Unassigned"
        else:
            claim_scheduled_dates[claim.id] = ""
            claim_assignments[claim.id] = "Unassigned"

    open_claims = [c for c in claims if claim_status_override[c.id] == "Open"]
    scheduled_claims = [c for c in claims if claim_status_override[c.id] == "Scheduled"]
    deferred_claims = [c for c in claims if claim_status_override[c.id] == "Deferred"]
    closed_claims = [c for c in claims if claim_status_override[c.id] == "Closed"]

    closures = {}
    deferrals = {}

    return render_template(
        'dashboard.html',
        open_claims=open_claims,
        scheduled_claims=scheduled_claims,
        deferred_claims=deferred_claims,
        closed_claims=closed_claims,
        claim_assignments=claim_assignments,
        claim_scheduled_dates=claim_scheduled_dates,
        claim_status_override=claim_status_override,
        closures=closures,
        deferrals=deferrals
    )

@app.route('/api/update_workorder_date', methods=['POST'])
@login_required
def update_workorder_date():
    from flask import jsonify  # make sure this import is present
    data = request.get_json()
    workorder_id = data.get('workorder_id')
    new_date = data.get('new_date')  # expected as "YYYY-MM-DD"
    try:
        workorder = WorkOrder.query.get_or_404(workorder_id)
        old_date = workorder.scheduled_date
        workorder.scheduled_date = datetime.strptime(new_date, "%Y-%m-%d").date()
        db.session.commit()
        
        # Log the change to ClaimLog
        db.session.add(ClaimLog(
            claim_id=workorder.claim_id,
            user_id=current_user.id,
            action=(
                f"Work order rescheduled from {old_date} to {workorder.scheduled_date} "
                f"by {current_user.name}."
            )
        ))
        db.session.commit()
        
        return jsonify({"success": True})
    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 500



@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        user = User.query.filter_by(email=request.form['email'].lower()).first()
        if user and check_password_hash(user.password, request.form['password']):
            login_user(user)
            return redirect(url_for('index'))
        else:
            flash('Invalid credentials')
    return render_template('login.html')

@app.route('/register', methods=['GET', 'POST'])
def register():
    if request.method == 'POST':
        email = request.form['email'].lower()
        name = request.form['name']
        password = request.form['password']
        if User.query.filter_by(email=email).first():
            flash('Email already registered. Please log in.')
            return redirect(url_for('login'))
        hashed = generate_password_hash(password)
        user = User(email=email, name=name, password=hashed)
        db.session.add(user)
        db.session.commit()
        flash('Registration successful! Please log in.')
        return redirect(url_for('login'))
    return render_template('register.html')

@app.route('/logout')
@login_required
def logout():
    logout_user()
    return redirect(url_for('login'))

@app.route('/add_vendor', methods=['GET', 'POST'])
@login_required
def add_vendor():
    if request.method == 'POST':
        name = request.form['name']
        contact_number = request.form['contact_number']
        email = request.form['email']
        vendor = Vendor(name=name, contact_number=contact_number, email=email)
        db.session.add(vendor)
        db.session.commit()
        flash('Vendor added!')
        return redirect(url_for('vendors'))
    return render_template('add_vendor.html')

@app.route('/vendors')
@login_required
def vendors():
    vendors = Vendor.query.all()
    return render_template('vendors.html', vendors=vendors)

@app.route('/add_assignee', methods=['GET', 'POST'])
@login_required
def add_assignee():
    if request.method == 'POST':
        name = request.form['name']
        contact_number = request.form['contact_number']
        email = request.form['email']
        assignee = Assignee(name=name, contact_number=contact_number, email=email)
        db.session.add(assignee)
        db.session.commit()
        flash('Assignee added!')
        return redirect(url_for('assignees'))
    return render_template('add_assignee.html')

@app.route('/assignees')
@login_required
def assignees():
    assignees = Assignee.query.all()
    return render_template('assignees.html', assignees=assignees)

@app.route('/add_claim', methods=['GET', 'POST'])
@login_required
def add_claim():
    if request.method == 'POST':
        try:
            address = request.form['address']
            homeowner_name = request.form['homeowner_name']
            homeowner_email = request.form['homeowner_email']
            homeowner_phone = request.form['homeowner_phone']
            cobuyer_name = request.form['cobuyer_name']
            cobuyer_email = request.form['cobuyer_email']
            cobuyer_phone = request.form['cobuyer_phone']
            warranty_type = request.form['warranty_type']
            issue_description = request.form['issue_description']

            claim = Claim(
                address=address,
                homeowner_name=homeowner_name,
                homeowner_email=homeowner_email,
                homeowner_phone=homeowner_phone,
                cobuyer_name=cobuyer_name,
                cobuyer_email=cobuyer_email,
                cobuyer_phone=cobuyer_phone,
                warranty_type=warranty_type,
                issue_description=issue_description
            )
            db.session.add(claim)
            db.session.commit()

            files = request.files.getlist('photos')
            for file in files:
                if file and file.filename:
                    filename = secure_filename(file.filename)
                    file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
                    photo = ClaimPhoto(claim_id=claim.id, filename=filename)
                    db.session.add(photo)
            db.session.commit()

            flash('Claim added!')
            return redirect(url_for('index'))
        except Exception as e:
            flash(f"Error adding claim: {e}")
    return render_template('add_claim.html')

@app.route('/view_claim/<int:claim_id>', methods=['GET', 'POST'])
@login_required
def view_claim(claim_id):
    claim = Claim.query.get_or_404(claim_id)
    vendors = Vendor.query.all()
    assignees = Assignee.query.all()
    workorders = WorkOrder.query.filter_by(claim_id=claim.id).order_by(WorkOrder.id.desc()).all()
    latest_workorder = workorders[0] if workorders else None

    if request.method == 'POST':
        vendor_id = request.form['vendor']
        assignee_id = request.form['assignee']
        scheduled_date = request.form['scheduled_date']
        scheduled_time = request.form['scheduled_time']
        notes = request.form['notes']

        vendor_id = int(vendor_id) if vendor_id else None
        assignee_id = int(assignee_id) if assignee_id else None

        workorder = WorkOrder(
            claim_id=claim.id,
            vendor_id=vendor_id,
            assignee_id=assignee_id,
            scheduled_date=datetime.strptime(scheduled_date, "%Y-%m-%d") if scheduled_date else None,
            scheduled_time=scheduled_time,
            status='Scheduled',
            notes=notes
        )
        db.session.add(workorder)
        claim.status = 'Scheduled'
        db.session.commit()
        flash('Workorder updated and claim scheduled!')
        return redirect(url_for('view_claim', claim_id=claim.id))

    return render_template('view_claim.html', claim=claim, vendors=vendors, assignees=assignees, latest_workorder=latest_workorder)

@app.route('/assign_workorder/<int:claim_id>', methods=['GET', 'POST'])
@login_required
def assign_workorder(claim_id):
    import pytz
    from ics import Calendar, Event
    import platform
    import subprocess
    from flask import url_for

    claim = Claim.query.get_or_404(claim_id)
    vendors = Vendor.query.all()
    assignees = Assignee.query.all()

    if request.method == 'POST':
        vendor_id = request.form['vendor']
        assignee_id = request.form['assignee']
        scheduled_date = request.form['scheduled_date']
        scheduled_time = request.form['scheduled_time']
        status = request.form['status']
        notes = request.form['notes']

        vendor_id = int(vendor_id) if vendor_id else None
        assignee_id = int(assignee_id) if assignee_id else None

        vendor = Vendor.query.get(vendor_id) if vendor_id else None
        assignee = Assignee.query.get(assignee_id) if assignee_id else None

        # Save WorkOrder
        workorder = WorkOrder(
            claim_id=claim_id,
            vendor_id=vendor_id,
            assignee_id=assignee_id,
            scheduled_date=datetime.strptime(scheduled_date, "%Y-%m-%d") if scheduled_date else None,
            scheduled_time=scheduled_time,
            status=status,
            notes=notes
        )
        db.session.add(workorder)
        claim.status = 'Scheduled'
        db.session.commit()
        flash('Work order assigned and claim set to Scheduled!')

        # LOG EVERY WORK ORDER ASSIGNMENT/CHANGE
        log_msg = (
            f"Work order assigned or changed: "
            f"Vendor: {vendor.name if vendor else 'N/A'}, "
            f"Assignee: {assignee.name if assignee else 'N/A'}, "
            f"Scheduled: {scheduled_date} {scheduled_time}, "
            f"Status: {status}, "
            f"Notes: {notes or 'N/A'} by {current_user.name}"
        )
        db.session.add(ClaimLog(
            claim_id=claim_id,
            user_id=current_user.id,
            action=log_msg
        ))
        db.session.commit()

        # --- LOG IF FIRST ASSIGNMENT (optional, doesn't hurt to keep) ---
        existing_wos = WorkOrder.query.filter_by(claim_id=claim_id).count()
        if existing_wos == 1:
            parts = []
            if assignee:
                parts.append(f"Assignee: {assignee.name}")
            if vendor:
                parts.append(f"Vendor: {vendor.name}")
            assignment_text = " and ".join(parts) if parts else "Unassigned"
            db.session.add(ClaimLog(
                claim_id=claim_id,
                user_id=current_user.id,
                action=f"Claim first assigned to {assignment_text} by {current_user.name}"
            ))
            db.session.commit()

        # --- Prepare blocks for the event description ---
        client_info = (
            "CLIENT CONTACT INFORMATION\n"
            "-------------------------\n"
            f"Name: {claim.homeowner_name or 'N/A'}\n"
            f"Email: {claim.homeowner_email or 'N/A'}\n"
            f"Phone: {claim.homeowner_phone or 'N/A'}\n"
            f"Cobuyer Name: {claim.cobuyer_name or 'N/A'}\n"
            f"Cobuyer Email: {claim.cobuyer_email or 'N/A'}\n"
            f"Cobuyer Phone: {claim.cobuyer_phone or 'N/A'}\n"
        )

        trade_info = (
            "TRADE OR VENDOR ASSIGNED TO WORK ORDER\n"
            "--------------------------------------\n"
            f"Vendor: {vendor.name if vendor else 'N/A'}\n"
            f"Vendor Email: {vendor.email if vendor else 'N/A'}\n"
            f"Vendor Phone: {vendor.contact_number if vendor else 'N/A'}\n"
            f"Assignee: {assignee.name if assignee else 'N/A'}\n"
            f"Assignee Email: {assignee.email if assignee else 'N/A'}\n"
            f"Assignee Phone: {assignee.contact_number if assignee else 'N/A'}\n"
        )

        # Photo links
        photo_links = []
        if hasattr(claim, "photos"):
            for photo in claim.photos:
                url = url_for('uploaded_file', filename=photo.filename, _external=True)
                photo_links.append(f"{photo.filename}: {url}")
        photos_section = "Photos:\n" + ("\n".join(photo_links) if photo_links else "None attached")

        claim_details = (
            "CLAIM DETAILS\n"
            "-------------\n"
            f"Address: {claim.address}\n"
            f"Warranty Type: {claim.warranty_type}\n"
            f"Issue: {claim.issue_description}\n"
            f"Notes: {notes or 'N/A'}\n"
            f"{photos_section}"
        )

        # Final event description, visually separated into sections
        description = (
            "==============================\n"
            f"{client_info}"
            "==============================\n"
            f"{trade_info}"
            "==============================\n"
            f"{claim_details}"
        )

        # Generate .ics calendar invite with CST timezone and real invitees
        try:
            central = pytz.timezone('America/Chicago')
            event_start_naive = datetime.strptime(f"{scheduled_date} {scheduled_time}", "%Y-%m-%d %H:%M")
            event_start = central.localize(event_start_naive)
            event_end = event_start + timedelta(hours=1)

            cal = Calendar()
            event = Event()
            event.name = f"Work Order for {claim.address}"
            event.begin = event_start
            event.end = event_end
            event.description = description
            cal.events.add(event)

            # --- Manual Attendee Injection ---
            ics_content = str(cal)
            attendee_lines = []
            if assignee and assignee.email:
                attendee_lines.append(f"ATTENDEE;CN={assignee.name}:mailto:{assignee.email}")
            if vendor and vendor.email:
                attendee_lines.append(f"ATTENDEE;CN={vendor.name}:mailto:{vendor.email}")

            # Insert attendee lines after DTSTART
            if attendee_lines:
                lines = ics_content.splitlines()
                new_lines = []
                inserted = False
                for line in lines:
                    new_lines.append(line)
                    if not inserted and line.startswith('DTSTART'):
                        new_lines.extend(attendee_lines)
                        inserted = True
                ics_content = "\r\n".join(new_lines) + "\r\n"

            # Save the .ics file
            ics_folder = os.path.join('static', 'ics')
            os.makedirs(ics_folder, exist_ok=True)
            ics_path = os.path.join(ics_folder, f"workorder_claim_{claim_id}.ics")
            with open(ics_path, 'w', encoding='utf-8') as f:
                f.write(ics_content)

            # Try to auto-open the .ics file (optional)
            abs_path = os.path.abspath(ics_path)
            if platform.system() == 'Windows':
                os.startfile(abs_path)
            elif platform.system() == 'Darwin':  # macOS
                subprocess.call(('open', abs_path))
            else:  # Linux
                subprocess.call(('xdg-open', abs_path))

        except Exception as e:
            flash(f"Failed to generate calendar invite: {e}", 'warning')

        return redirect(url_for('index'))

    return render_template('assign_workorder.html', claim=claim, vendors=vendors, assignees=assignees)


@app.route('/update_claim_status/<int:claim_id>', methods=['POST'])
@login_required
def update_claim_status(claim_id):
    claim = Claim.query.get_or_404(claim_id)
    old_status = claim.status
    new_status = request.form.get('status')
    if new_status != old_status:
        if new_status == "Deferred":
            return redirect(url_for('defer_claim', claim_id=claim_id))
        if new_status == "Closed":
            return redirect(url_for('close_claim', claim_id=claim_id))
        claim.status = new_status
        db.session.add(ClaimLog(
            claim_id=claim.id,
            user_id=current_user.id,
            action=f"Status changed from {old_status} to {new_status}"
        ))
        db.session.commit()
        flash('Claim status updated and action logged.')
    return redirect(url_for('index'))

@app.route('/defer_claim/<int:claim_id>', methods=['GET', 'POST'])
@login_required
def defer_claim(claim_id):
    claim = Claim.query.get_or_404(claim_id)
    if request.method == 'POST':
        reason = request.form.get('reason', '')
        old_status = claim.status
        claim.status = "Deferred"
        db.session.add(ClaimLog(
            claim_id=claim.id,
            user_id=current_user.id,
            action=f"Status changed from {old_status} to Deferred. Reason: {reason}"
        ))
        db.session.commit()
        flash('Claim deferred and action logged.')
        return redirect(url_for('index'))
    return render_template('defer_claim.html', claim=claim)


@app.route('/delete_claim/<int:claim_id>', methods=['POST'])
@login_required
def delete_claim(claim_id):
    claim = Claim.query.get_or_404(claim_id)
    WorkOrder.query.filter_by(claim_id=claim_id).delete()
    ClaimLog.query.filter_by(claim_id=claim_id).delete()
    for photo in claim.photos:
        try:
            os.remove(os.path.join(app.config['UPLOAD_FOLDER'], photo.filename))
        except Exception:
            pass
    ClaimPhoto.query.filter_by(claim_id=claim_id).delete()
    db.session.delete(claim)
    db.session.commit()
    flash('Claim deleted successfully.')
    return redirect(url_for('index'))


@app.route('/claim_log/<int:claim_id>')
@login_required
def view_claim_log(claim_id):
    logs = ClaimLog.query.filter_by(claim_id=claim_id).order_by(ClaimLog.timestamp.desc()).all()
    return render_template('claim_log.html', logs=logs, claim_id=claim_id)

@app.route('/static/uploads/<filename>')
def uploaded_file(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename)

@app.route('/calendar')
@login_required
def calendar_view():
    now = datetime.now()
    year = int(request.args.get('year', now.year))
    month = int(request.args.get('month', now.month))

    start_date = datetime(year, month, 1).date()
    if month == 12:
        end_date = datetime(year + 1, 1, 1).date()
    else:
        end_date = datetime(year, month + 1, 1).date()

    workorders = (
        WorkOrder.query
        .join(Claim, WorkOrder.claim_id == Claim.id)
        .filter(
            WorkOrder.scheduled_date >= start_date,
            WorkOrder.scheduled_date < end_date,
            Claim.status != "Closed"
        )
        .all()
    )

    # Only show the workorder with the LATEST scheduled_date for each claim
    latest_workorder_per_claim = {}
    for wo in workorders:
        if wo.claim_id not in latest_workorder_per_claim:
            latest_workorder_per_claim[wo.claim_id] = wo
        else:
            if wo.scheduled_date and wo.scheduled_date > latest_workorder_per_claim[wo.claim_id].scheduled_date:
                latest_workorder_per_claim[wo.claim_id] = wo

    workorders_by_day = {}
    for wo in latest_workorder_per_claim.values():
        if wo.scheduled_date:
            day = wo.scheduled_date.day
            wo_dict = {
                'id': wo.id,
                'claim_id': wo.claim_id,
                'vendor': {'name': wo.vendor.name} if wo.vendor else None,
                'assignee': {'name': wo.assignee.name} if wo.assignee else None,
                'scheduled_date': wo.scheduled_date.strftime('%Y-%m-%d'),
                'claim': {'address': wo.claim.address if wo.claim else ''}
            }
            workorders_by_day.setdefault(day, []).append(wo_dict)

    cal = calendar.monthcalendar(year, month)
    month_name = calendar.month_name[month]

    return render_template(
        'calendar.html',
        cal=cal,
        workorders_by_day=workorders_by_day,
        month=month,
        year=year,
        month_name=month_name,
        now=now
    )

# ... (defer_claim, update_claim_status, claim_log, delete_claim, close_claim, file serving, api, etc. unchanged)

if __name__ == '__main__':
    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
    os.makedirs(app.config['ICS_FOLDER'], exist_ok=True)
    with app.app_context():
        db.create_all()
    app.run(debug=True)
