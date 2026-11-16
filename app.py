import os
import csv
import re
import calendar as cal
import secrets
import hashlib
import smtplib
from email.message import EmailMessage
from datetime import date, datetime, timedelta
from functools import wraps
from urllib.parse import urlparse, parse_qs
from urllib.request import Request, urlopen

from flask import Flask, render_template, request, redirect, url_for, flash, jsonify, send_file, session
from flask_sqlalchemy import SQLAlchemy
from flask_login import (
    LoginManager, UserMixin, login_user, logout_user, login_required, current_user
)
from werkzeug.security import generate_password_hash, check_password_hash
from sqlalchemy import UniqueConstraint, inspect
from io import BytesIO, StringIO
from openpyxl import Workbook, load_workbook
from openpyxl.utils.datetime import from_excel
import xlrd
from docx import Document

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
app = Flask(__name__)
app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", "dev-change-this-secret-key")
# Use PostgreSQL when DATABASE_URL is configured (Render); otherwise use local SQLite.
database_url = os.environ.get("DATABASE_URL")
if database_url:
    # Render and some providers may supply the legacy postgres:// scheme.
    if database_url.startswith("postgres://"):
        database_url = database_url.replace("postgres://", "postgresql+psycopg://", 1)
    elif database_url.startswith("postgresql://"):
        database_url = database_url.replace("postgresql://", "postgresql+psycopg://", 1)
    app.config["SQLALCHEMY_DATABASE_URI"] = database_url
else:
    app.config["SQLALCHEMY_DATABASE_URI"] = "sqlite:///" + os.path.join(BASE_DIR, "tracker.db")
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"
app.config["MAX_CONTENT_LENGTH"] = 12 * 1024 * 1024

db = SQLAlchemy(app)
login_manager = LoginManager(app)
login_manager.login_view = "login"
login_manager.login_message_category = "warning"


# ----------------------------- DATABASE MODELS -----------------------------
class User(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(50), unique=True, nullable=False)
    full_name = db.Column(db.String(100), nullable=False)
    password_hash = db.Column(db.String(255), nullable=False)
    role = db.Column(db.String(10), nullable=False)  # admin, tl, agent
    tl_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=True)
    active = db.Column(db.Boolean, default=True, nullable=False)
    email = db.Column(db.String(254), unique=True, nullable=True)
    email_verified = db.Column(db.Boolean, default=False, nullable=False)

    team_lead = db.relationship("User", remote_side=[id], backref="team_members", uselist=False)

    def set_password(self, password):
        self.password_hash = generate_password_hash(password)

    def check_password(self, password):
        return check_password_hash(self.password_hash, password)


class EmailVerificationCode(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)
    email = db.Column(db.String(254), nullable=False)
    purpose = db.Column(db.String(30), nullable=False)
    code_hash = db.Column(db.String(64), nullable=False)
    expires_at = db.Column(db.DateTime, nullable=False)
    attempts = db.Column(db.Integer, nullable=False, default=0)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)


class DailyEntry(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    agent_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)
    entry_date = db.Column(db.Date, nullable=False)
    total_calls = db.Column(db.Integer, default=0, nullable=False)
    approvals = db.Column(db.Integer, default=0, nullable=False)
    errors = db.Column(db.Integer, default=0, nullable=False)
    __table_args__ = (UniqueConstraint("agent_id", "entry_date", name="uq_agent_entry_date"),)

    agent = db.relationship("User", backref="daily_entries")

    @property
    def approval_pct(self):
        return round(self.approvals * 100 / self.total_calls, 2) if self.total_calls else 0

    @property
    def error_pct(self):
        return round(self.errors * 100 / self.approvals, 2) if self.approvals else 0


class EntryHistory(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    entry_id = db.Column(db.Integer, db.ForeignKey("daily_entry.id"), nullable=False)
    changed_by = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)
    changed_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    old_values = db.Column(db.String(200))
    new_values = db.Column(db.String(200))
    editor = db.relationship("User")


class Roster(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    agent_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)
    off_date = db.Column(db.Date, nullable=False)
    roster_type = db.Column(db.String(20), nullable=False, default="weekoff")
    agent_entered = db.Column(db.Boolean, nullable=False, default=False)
    agent = db.relationship("User", backref="roster_days")
    __table_args__ = (UniqueConstraint("agent_id", "off_date", name="uq_agent_off_date"),)


class ErrorSpreadsheet(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    tl_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)
    month = db.Column(db.String(7), nullable=False)
    spreadsheet_url = db.Column(db.String(1000), nullable=False, default="")
    source_name = db.Column(db.String(255))
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow, nullable=False)
    team_lead = db.relationship("User")
    __table_args__ = (UniqueConstraint("tl_id", "month", name="uq_tl_error_spreadsheet_month"),)


class AgentError(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    source_id = db.Column(db.Integer, db.ForeignKey("error_spreadsheet.id"), nullable=False)
    tl_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)
    agent_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)
    month = db.Column(db.String(7), nullable=False)
    error_id = db.Column(db.String(100), nullable=False)
    reason = db.Column(db.Text, nullable=False)
    error_date = db.Column(db.Date, nullable=True)
    agent_name = db.Column(db.String(150), nullable=True)
    category = db.Column(db.Text, nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    agent = db.relationship("User", foreign_keys=[agent_id])
    team_lead = db.relationship("User", foreign_keys=[tl_id])
    source = db.relationship("ErrorSpreadsheet", backref=db.backref("errors", cascade="all, delete-orphan"))
    __table_args__ = (UniqueConstraint("source_id", "agent_id", "error_id", name="uq_source_agent_error"),)


class Target(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    month = db.Column(db.String(7), unique=True, nullable=False)
    avg_calls = db.Column(db.Float, default=170)
    approval_pct = db.Column(db.Float, default=45)
    error_pct = db.Column(db.Float, default=0.5)


@login_manager.user_loader
def load_user(uid):
    return db.session.get(User, int(uid))


# -------------------------------- HELPERS ----------------------------------
def role_required(*roles):
    def decorator(fn):
        @wraps(fn)
        def wrapper(*args, **kwargs):
            if not current_user.is_authenticated or current_user.role not in roles:
                return "Access denied", 403
            return fn(*args, **kwargs)
        return wrapper
    return decorator


def get_target(month):
    target = Target.query.filter_by(month=month).first()
    if not target:
        target = Target(month=month, avg_calls=170, approval_pct=45, error_pct=0.5)
        db.session.add(target)
        db.session.commit()
    return target


def month_bounds(year, month):
    return date(year, month, 1), date(year, month, cal.monthrange(year, month)[1])


def my_agents():
    return User.query.filter_by(role="agent", tl_id=current_user.id, active=True).order_by(User.full_name).all()


def working_days_between(agent_id, start_date, end_date):
    """Count entered work dates, excluding only dates the agent marked Week-off or Leave."""
    if end_date < start_date:
        return []

    entered_dates = {
        entry_date for (entry_date,) in db.session.query(DailyEntry.entry_date).filter(
            DailyEntry.agent_id == agent_id,
            DailyEntry.entry_date >= start_date,
            DailyEntry.entry_date <= end_date
        ).distinct().all()
    }
    marked_off_dates = {
        off_date for (off_date,) in db.session.query(Roster.off_date).filter(
            Roster.agent_id == agent_id,
            Roster.off_date >= start_date,
            Roster.off_date <= end_date,
            Roster.roster_type.in_(("weekoff", "leave")),
            Roster.agent_entered.is_(True)
        ).distinct().all()
    }
    return sorted(entered_dates - marked_off_dates)


def selected_period():
    """Return a validated date range; all dashboard ranges stop at today."""
    today = date.today()
    period = request.args.get("period", "month")
    try:
        if period == "day":
            chosen = datetime.strptime(request.args.get("date", today.isoformat()), "%Y-%m-%d").date()
            start = end = chosen
        elif period == "week":
            chosen = datetime.strptime(request.args.get("date", today.isoformat()), "%Y-%m-%d").date()
            start = date.fromordinal(chosen.toordinal() - chosen.weekday())
            end = date.fromordinal(start.toordinal() + 6)
        elif period == "custom":
            start = datetime.strptime(request.args["start"], "%Y-%m-%d").date()
            end = datetime.strptime(request.args["end"], "%Y-%m-%d").date()
            if end < start:
                raise ValueError
        else:
            period = "month"
            year = int(request.args.get("y", today.year))
            month = int(request.args.get("m", today.month))
            start, end = month_bounds(year, month)
        if start > today:
            start = today
        end = min(end, today)
        if end < start:
            end = start
        return period, start, end
    except (KeyError, TypeError, ValueError):
        first, last = month_bounds(today.year, today.month)
        return "month", first, today


def agent_kpis(agent_id, start, end):
    month_target = get_target(start.strftime("%Y-%m"))
    entries = DailyEntry.query.filter(
        DailyEntry.agent_id == agent_id,
        DailyEntry.entry_date >= start,
        DailyEntry.entry_date <= end
    ).order_by(DailyEntry.entry_date).all()
    total_calls = sum(e.total_calls for e in entries)
    total_approvals = sum(e.approvals for e in entries)
    total_errors = sum(e.errors for e in entries)
    work_days = working_days_between(agent_id, start, end)
    divisor = len(work_days)
    avg_calls = round(total_calls / divisor, 1) if divisor else 0
    avg_approvals = round(total_approvals / divisor, 1) if divisor else 0
    approval_pct = round(total_approvals * 100 / total_calls, 2) if total_calls else 0
    error_pct = round(total_errors * 100 / total_approvals, 2) if total_approvals else 0
    chart = {
        "labels": [e.entry_date.strftime("%d %b") for e in entries],
        "calls": [e.total_calls for e in entries],
        "approvals": [e.approvals for e in entries],
        "approval_pct": [e.approval_pct for e in entries],
        "error_pct": [e.error_pct for e in entries],
    }
    return {
        "target": month_target, "days_done": len(entries), "working_days": divisor,
        "total_calls": total_calls, "total_appr": total_approvals, "total_err": total_errors,
        "avg_calls": avg_calls, "avg_approvals": avg_approvals,
        "appr_pct": approval_pct, "err_pct": error_pct,
        "reminder": [], "chart": chart, "entries": entries
    }


# -------------------------------- AUTH --------------------------------------
def send_verification_email(email, code, purpose):
    mail_user = os.environ.get("MAIL_USERNAME")
    mail_password = os.environ.get("MAIL_PASSWORD")
    if not mail_user or not mail_password:
        raise RuntimeError("Email sending is not configured. Set MAIL_USERNAME and MAIL_PASSWORD.")
    message = EmailMessage()
    message["Subject"] = "Your Daily Tracker verification code"
    message["From"] = os.environ.get("MAIL_DEFAULT_SENDER", mail_user)
    message["To"] = email
    action = "verify your company email address" if purpose == "link_email" else "reset your Daily Tracker password"
    message.set_content(
        f"Hello,\n\nYour code to {action} is: {code}\n\n"
        "This code expires in 10 minutes. Do not share it with anyone.\n\nDaily Performance Tracker"
    )
    server = os.environ.get("MAIL_SERVER", "smtp.gmail.com")
    port = int(os.environ.get("MAIL_PORT", "587"))
    with smtplib.SMTP(server, port, timeout=20) as smtp:
        smtp.starttls()
        smtp.login(mail_user, mail_password)
        smtp.send_message(message)


def issue_email_code(user, email, purpose):
    EmailVerificationCode.query.filter_by(user_id=user.id, purpose=purpose).delete()
    code = f"{secrets.randbelow(1000000):06d}"
    record = EmailVerificationCode(
        user_id=user.id, email=email, purpose=purpose,
        code_hash=hashlib.sha256(code.encode()).hexdigest(),
        expires_at=datetime.utcnow() + timedelta(minutes=10), attempts=0
    )
    db.session.add(record)
    db.session.commit()
    try:
        send_verification_email(email, code, purpose)
    except Exception:
        db.session.delete(record)
        db.session.commit()
        raise


@app.route("/")
def home():
    if not current_user.is_authenticated:
        return redirect(url_for("login"))
    if current_user.role == "admin":
        return redirect(url_for("admin_users"))
    if current_user.role == "tl":
        return redirect(url_for("tl_team"))
    return redirect(url_for("agent_dashboard"))


@app.route("/login", methods=["GET", "POST"])
def login():
    if current_user.is_authenticated:
        return redirect(url_for("home"))
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        user = User.query.filter_by(username=username, active=True).first()
        if user and user.check_password(password):
            login_user(user)
            return redirect(url_for("home"))
        flash("Incorrect username or password.", "danger")
    return render_template("login.html")


@app.route("/agent/email-verification", methods=["POST"])
@login_required
@role_required("agent")
def agent_email_verification():
    action = request.form.get("action")
    if action == "send":
        email = request.form.get("email", "").strip().lower()
        if not re.fullmatch(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}", email):
            flash("Enter a valid company email address.", "danger")
        else:
            owner = User.query.filter(User.email == email, User.id != current_user.id).first()
            if owner:
                flash("That email address is already linked to another account.", "danger")
            else:
                try:
                    issue_email_code(current_user, email, "link_email")
                    session["pending_agent_email"] = email
                    flash("Verification code sent. Enter it below to verify your company email.", "success")
                except Exception:
                    app.logger.exception("Could not send company email verification")
                    flash("Email could not be sent. Please contact your administrator.", "danger")
    elif action == "verify":
        email = session.get("pending_agent_email")
        code = request.form.get("code", "").strip()
        record = EmailVerificationCode.query.filter_by(
            user_id=current_user.id, purpose="link_email", email=email
        ).order_by(EmailVerificationCode.created_at.desc()).first() if email else None
        if not record or record.expires_at < datetime.utcnow() or record.attempts >= 5:
            flash("The code has expired. Request a new code.", "danger")
        else:
            record.attempts += 1
            expected = record.code_hash
            actual = hashlib.sha256(code.encode()).hexdigest()
            if secrets.compare_digest(expected, actual):
                current_user.email = email
                current_user.email_verified = True
                db.session.delete(record)
                db.session.commit()
                session.pop("pending_agent_email", None)
                flash("Company email verified successfully. You can now use Forgot Password.", "success")
            else:
                db.session.commit()
                flash("Incorrect verification code. Please try again.", "danger")
    return redirect(url_for("home"))


@app.route("/forgot-password", methods=["GET", "POST"])
def forgot_password():
    if request.method == "POST":
        email = request.form.get("email", "").strip().lower()
        session.pop("password_reset_user_id", None)
        user = User.query.filter_by(email=email, email_verified=True, role="agent", active=True).first()
        if user:
            try:
                issue_email_code(user, email, "reset_password")
                session["password_reset_user_id"] = user.id
                session["password_reset_email"] = email
            except Exception:
                app.logger.exception("Could not send password reset email")
        flash("If an active agent account has this verified email address, a reset code has been sent.", "success")
        return redirect(url_for("reset_password"))
    return render_template("forgot_password.html")


@app.route("/reset-password", methods=["GET", "POST"])
def reset_password():
    user_id = session.get("password_reset_user_id")
    email = session.get("password_reset_email")
    if request.method == "POST":
        user = db.session.get(User, user_id) if user_id else None
        code = request.form.get("code", "").strip()
        password = request.form.get("new_password", "")
        confirm = request.form.get("confirm_password", "")
        record = EmailVerificationCode.query.filter_by(
            user_id=user_id, email=email, purpose="reset_password"
        ).order_by(EmailVerificationCode.created_at.desc()).first() if user_id else None
        if not user or not record or record.expires_at < datetime.utcnow() or record.attempts >= 5:
            flash("The reset code is invalid or expired. Please start again.", "danger")
        elif len(password) < 8 or password != confirm:
            flash("Passwords must match and contain at least 8 characters.", "danger")
        else:
            record.attempts += 1
            if secrets.compare_digest(record.code_hash, hashlib.sha256(code.encode()).hexdigest()):
                user.set_password(password)
                db.session.delete(record)
                db.session.commit()
                session.pop("password_reset_user_id", None)
                session.pop("password_reset_email", None)
                flash("Password reset successfully. Please sign in.", "success")
                return redirect(url_for("login"))
            db.session.commit()
            flash("Incorrect verification code.", "danger")
    return render_template("reset_password.html", reset_email=email)


@app.route("/logout", methods=["POST"])
@login_required
def logout():
    logout_user()
    flash("You have been logged out.", "success")
    return redirect(url_for("login"))


@app.route("/change-password", methods=["POST"])
@login_required
def change_password():
    new_password = request.form.get("new_password", "")
    if len(new_password) < 8:
        flash("Password must contain at least 8 characters.", "danger")
    else:
        current_user.set_password(new_password)
        db.session.commit()
        flash("Password updated.", "success")
    return redirect(url_for("home"))


# -------------------------------- ADMIN -------------------------------------
@app.route("/admin/users", methods=["GET", "POST"])
@login_required
@role_required("admin")
def admin_users():
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        full_name = request.form.get("full_name", "").strip()
        password = request.form.get("password", "")
        role = request.form.get("role", "")
        tl_id = request.form.get("tl_id") or None
        if not username or not full_name or len(password) < 8 or role not in ("admin", "tl", "agent"):
            flash("Enter valid details. Password must be at least 8 characters.", "danger")
        elif User.query.filter_by(username=username).first():
            flash("Username already exists.", "danger")
        else:
            if role == "agent":
                try:
                    tl_id = int(tl_id) if tl_id else None
                except ValueError:
                    tl_id = None
                if not tl_id or not db.session.get(User, tl_id) or db.session.get(User, tl_id).role != "tl":
                    flash("Select a valid Team Lead for the agent.", "danger")
                    return redirect(url_for("admin_users"))
            else:
                tl_id = None
            user = User(username=username, full_name=full_name, role=role, tl_id=tl_id)
            user.set_password(password)
            db.session.add(user)
            db.session.commit()
            flash("User created successfully.", "success")
        return redirect(url_for("admin_users"))

    users = User.query.order_by(User.role, User.full_name).all()
    tls = User.query.filter_by(role="tl", active=True).order_by(User.full_name).all()
    target = get_target(date.today().strftime("%Y-%m"))
    return render_template("admin_users.html", users=users, tls=tls, target=target,
                           month=date.today().strftime("%Y-%m"))


@app.route("/admin/user/<int:user_id>/update", methods=["POST"])
@login_required
@role_required("admin")
def update_user(user_id):
    user = db.session.get(User, user_id)
    if not user or user.id == current_user.id:
        flash("You cannot edit this administrator account here.", "danger")
        return redirect(url_for("admin_users"))

    username = request.form.get("username", "").strip()
    full_name = request.form.get("full_name", "").strip()
    role = request.form.get("role", "")
    tl_raw = request.form.get("tl_id") or None

    if not username or not full_name or role not in ("admin", "tl", "agent"):
        flash("Enter a valid name, username, and role.", "danger")
        return redirect(url_for("admin_users"))
    duplicate = User.query.filter(User.username == username, User.id != user.id).first()
    if duplicate:
        flash("That username is already in use.", "danger")
        return redirect(url_for("admin_users"))
    if user.team_members and role != "tl":
        flash("This Team Lead still has assigned agents. Reassign those agents before changing the role.", "danger")
        return redirect(url_for("admin_users"))

    tl_id = None
    if role == "agent":
        try:
            tl_id = int(tl_raw) if tl_raw else None
        except (TypeError, ValueError):
            tl_id = None
        tl = db.session.get(User, tl_id) if tl_id else None
        if not tl or tl.role != "tl" or not tl.active:
            flash("Choose an active Team Lead for this agent.", "danger")
            return redirect(url_for("admin_users"))

    user.username = username
    user.full_name = full_name
    user.role = role
    user.tl_id = tl_id
    db.session.commit()
    flash("User details updated.", "success")
    return redirect(url_for("admin_users"))


@app.route("/admin/user/<int:user_id>/reset-password", methods=["POST"])
@login_required
@role_required("admin")
def admin_reset_password(user_id):
    user = db.session.get(User, user_id)
    new_password = request.form.get("new_password", "")
    if not user or user.id == current_user.id:
        flash("You cannot reset this administrator password here.", "danger")
    elif len(new_password) < 8:
        flash("The new password must be at least 8 characters.", "danger")
    else:
        user.set_password(new_password)
        db.session.commit()
        flash(f"Password reset for {user.full_name}. Share the new password securely.", "success")
    return redirect(url_for("admin_users"))


@app.route("/admin/user/<int:user_id>/toggle", methods=["POST"])
@login_required
@role_required("admin")
def toggle_user(user_id):
    user = db.session.get(User, user_id)
    if not user or user.id == current_user.id:
        flash("That account cannot be changed.", "danger")
    else:
        user.active = not user.active
        db.session.commit()
        flash("User status updated.", "success")
    return redirect(url_for("admin_users"))


@app.route("/admin/targets", methods=["POST"])
@login_required
@role_required("admin")
def admin_targets():
    try:
        month = request.form["month"]
        datetime.strptime(month, "%Y-%m")
        avg_calls = float(request.form["avg_calls"])
        approval_pct = float(request.form["approval_pct"])
        error_pct = float(request.form["error_pct"])
        if avg_calls < 0 or not 0 <= approval_pct <= 100 or not 0 <= error_pct <= 100:
            raise ValueError
        target = get_target(month)
        target.avg_calls, target.approval_pct, target.error_pct = avg_calls, approval_pct, error_pct
        db.session.commit()
        flash("Monthly targets saved.", "success")
    except (KeyError, ValueError):
        flash("Please enter valid target values.", "danger")
    return redirect(url_for("admin_users"))


@app.route("/admin/roster", methods=["GET", "POST"])
@login_required
@role_required("admin")
def admin_roster():
    agents = User.query.filter_by(role="agent", active=True).order_by(User.full_name).all()
    if request.method == "POST":
        try:
            agent_id = int(request.form["agent_id"])
            off_date = datetime.strptime(request.form["off_date"], "%Y-%m-%d").date()
            agent = db.session.get(User, agent_id)
            if not agent or agent.role != "agent":
                raise ValueError
            existing = Roster.query.filter_by(agent_id=agent_id, off_date=off_date).first()
            if not existing:
                db.session.add(Roster(agent_id=agent_id, off_date=off_date))
                db.session.commit()
            flash("Week-off date saved.", "success")
        except (KeyError, ValueError):
            flash("Please select a valid agent and date.", "danger")
        return redirect(url_for("admin_roster"))
    roster = Roster.query.order_by(Roster.off_date.desc()).limit(100).all()
    return render_template("admin_roster.html", agents=agents, roster=roster)


@app.route("/admin/roster/<int:roster_id>/delete", methods=["POST"])
@login_required
@role_required("admin")
def delete_roster(roster_id):
    row = db.session.get(Roster, roster_id)
    if row:
        db.session.delete(row)
        db.session.commit()
        flash("Week-off removed.", "success")
    return redirect(url_for("admin_roster"))

def parse_month():
    today = date.today()

    try:
        year = int(request.args.get("y", today.year))
        month = int(request.args.get("m", today.month))

        if not 2000 <= year <= 2100 or not 1 <= month <= 12:
            raise ValueError

        return year, month

    except (ValueError, TypeError):
        return today.year, today.month
# -------------------------------- AGENT -------------------------------------
@app.route("/agent")
@login_required
@role_required("agent")
def agent_dashboard():
    period, start, end = selected_period()
    k = agent_kpis(current_user.id, start, end)
    return render_template("agent_dashboard.html", k=k, y=start.year, m=start.month,
                           period=period, start=start, end=end,
                           month_name=cal.month_name[start.month])


@app.route("/agent/roster", methods=["GET", "POST"])
@login_required
@role_required("agent")
def agent_roster():
    if request.method == "POST":
        try:
            off_date = datetime.strptime(
                request.form["off_date"], "%Y-%m-%d"
            ).date()
            roster_type = request.form.get("roster_type", "").strip().lower()
            if roster_type not in ("weekoff", "leave"):
                raise ValueError

            existing = Roster.query.filter_by(
                agent_id=current_user.id,
                off_date=off_date
            ).first()

            if existing:
                existing.roster_type = roster_type
                existing.agent_entered = True
            else:
                db.session.add(
                    Roster(
                        agent_id=current_user.id,
                        off_date=off_date,
                        roster_type=roster_type,
                        agent_entered=True
                    )
                )
            db.session.commit()
            label = "Week-off" if roster_type == "weekoff" else "Leave"
            flash(f"Your {label.lower()} date has been saved.", "success")

        except (KeyError, ValueError):
            db.session.rollback()
            flash("Choose a valid date and select Week-off or Leave.", "danger")

        return redirect(url_for("agent_roster"))

    roster = Roster.query.filter_by(
        agent_id=current_user.id
    ).order_by(Roster.off_date.asc()).all()

    return render_template(
        "agent_roster.html",
        roster=roster,
        today=date.today()
    )


@app.route("/agent/roster/<int:roster_id>/delete", methods=["POST"])
@login_required
@role_required("agent")
def agent_delete_roster(roster_id):
    row = db.session.get(Roster, roster_id)
    if not row or row.agent_id != current_user.id:
        flash("That week-off date does not belong to your account.", "danger")
    else:
        db.session.delete(row)
        db.session.commit()
        flash("Your week-off date has been removed.", "success")
    return redirect(url_for("agent_roster"))


@app.route("/agent/calendar")
@login_required
@role_required("agent")
def agent_calendar():
    year, month = parse_month()
    first, last = month_bounds(year, month)
    entries = {e.entry_date.day: e for e in DailyEntry.query.filter(
        DailyEntry.agent_id == current_user.id,
        DailyEntry.entry_date >= first, DailyEntry.entry_date <= last).all()}
    roster_rows = Roster.query.filter(
        Roster.agent_id == current_user.id, Roster.off_date >= first, Roster.off_date <= last
    ).all()
    offs = {r.off_date.day for r in roster_rows}
    roster_types = {r.off_date.day: r.roster_type for r in roster_rows}
    entry_data = {str(day): {"total_calls": e.total_calls, "approvals": e.approvals, "errors": e.errors}
                  for day, e in entries.items()}
    weeks = cal.Calendar(firstweekday=0).monthdayscalendar(year, month)
    return render_template("calendar.html", y=year, m=month, weeks=weeks, entries=entries,
                           entry_data=entry_data, offs=offs, roster_types=roster_types,
                           month_name=cal.month_name[month], today=date.today())


@app.route("/agent/entry", methods=["POST"])
@login_required
@role_required("agent")
def save_entry():
    try:
        entry_date = datetime.strptime(request.form["entry_date"], "%Y-%m-%d").date()
        calls = int(request.form["total_calls"])
        approvals = int(request.form["approvals"])
        errors = int(request.form["errors"])
        if min(calls, approvals, errors) < 0 or approvals > calls or errors > calls:
            raise ValueError
        if entry_date > date.today():
            flash("You cannot enter performance for a future date.", "danger")
            return redirect(url_for("agent_calendar", y=entry_date.year, m=entry_date.month))
        entry = DailyEntry.query.filter_by(agent_id=current_user.id, entry_date=entry_date).first()
        if entry:
            old_values = f"{entry.total_calls}/{entry.approvals}/{entry.errors}"
            entry.total_calls, entry.approvals, entry.errors = calls, approvals, errors
        else:
            old_values = "new entry"
            entry = DailyEntry(agent_id=current_user.id, entry_date=entry_date,
                               total_calls=calls, approvals=approvals, errors=errors)
            db.session.add(entry)
            db.session.flush()
        db.session.add(EntryHistory(entry_id=entry.id, changed_by=current_user.id,
                                    old_values=old_values,
                                    new_values=f"{calls}/{approvals}/{errors}"))
        db.session.commit()
        flash("Daily performance saved.", "success")
        return redirect(url_for("agent_calendar", y=entry_date.year, m=entry_date.month))
    except (KeyError, ValueError):
        flash("Please enter valid whole-number values. Approvals and errors cannot exceed calls.", "danger")
        return redirect(url_for("agent_calendar"))


# -------------------------- ERROR MANAGEMENT -------------------------------
def normalize_error_header(value):
    return re.sub(r"[^a-z0-9]", "", str(value or "").strip().lower())


def parse_error_date(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)):
        try:
            return from_excel(value).date()
        except Exception:
            try:
                return xlrd.xldate_as_datetime(value, 0).date()
            except Exception:
                pass
    text_value = str(value or "").strip()
    if not text_value:
        return None
    for fmt in (
        "%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y",
        "%Y/%m/%d", "%d.%m.%Y", "%Y-%m-%d %H:%M:%S",
        "%d/%m/%Y %H:%M:%S", "%d-%b-%Y", "%d %b %Y"
    ):
        try:
            return datetime.strptime(text_value, fmt).date()
        except ValueError:
            continue
    try:
        return datetime.fromisoformat(text_value.replace("Z", "+00:00")).date()
    except ValueError:
        raise ValueError("Could not read a Date value. Use a normal Excel date or DD/MM/YYYY / YYYY-MM-DD.")


def google_sheet_csv_url(sheet_url):
    parsed = urlparse(sheet_url)
    if parsed.scheme != "https" or parsed.netloc.lower() not in ("docs.google.com", "www.docs.google.com"):
        raise ValueError("Use a Google Sheets URL or upload an Excel, CSV, or DOCX file.")
    match = re.search(r"/spreadsheets/d/([a-zA-Z0-9_-]+)", parsed.path)
    if not match:
        raise ValueError("The Google Sheets URL is not valid.")
    sheet_id = match.group(1)
    gid = parse_qs(parsed.query).get("gid", ["0"])[0]
    return f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}"


def rows_from_matrix(matrix):
    matrix = [list(row) for row in matrix]
    while matrix and not any(str(cell or "").strip() for cell in matrix[0]):
        matrix.pop(0)
    if not matrix:
        raise ValueError("The spreadsheet or document is empty.")
    headers = [str(value or "").strip() for value in matrix[0]]
    return [
        {headers[i]: row[i] if i < len(row) else None for i in range(len(headers))}
        for row in matrix[1:]
    ]


def read_error_rows(file_bytes, filename, month):
    extension = os.path.splitext(filename.lower())[1]
    if extension == ".csv":
        text_data = file_bytes.decode("utf-8-sig")
        reader = csv.reader(StringIO(text_data))
        raw_rows = rows_from_matrix(reader)
    elif extension in (".xlsx", ".xlsm"):
        workbook = load_workbook(BytesIO(file_bytes), read_only=True, data_only=True)
        worksheet = workbook.active
        raw_rows = rows_from_matrix(worksheet.iter_rows(values_only=True))
        workbook.close()
    elif extension == ".xls":
        workbook = xlrd.open_workbook(file_contents=file_bytes)
        worksheet = workbook.sheet_by_index(0)
        raw_rows = rows_from_matrix(
            worksheet.row_values(row_index) for row_index in range(worksheet.nrows)
        )
    elif extension == ".docx":
        document = Document(BytesIO(file_bytes))
        if not document.tables:
            raise ValueError("The DOCX file must contain a table with the required column headers.")
        table = document.tables[0]
        raw_rows = rows_from_matrix(
            [[cell.text.strip() for cell in row.cells] for row in table.rows]
        )
    else:
        raise ValueError("Supported uploads: .xlsx, .xlsm, .xls, .csv, and .docx files.")

    if not raw_rows:
        raise ValueError("No data rows were found in the file.")

    headers = {normalize_error_header(key): key for key in raw_rows[0].keys()}
    date_key = next((headers[key] for key in ("date", "errordate", "createddate") if key in headers), None)
    username_key = next((headers[key] for key in (
        "empid", "employeeid", "employeeusername", "username", "agentusername", "agentuser", "login", "agentlogin"
    ) if key in headers), None)
    name_key = next((headers[key] for key in ("agentname", "employeename", "fullname", "name") if key in headers), None)
    uuid_key = next((headers[key] for key in (
        "uuid", "errorid", "erroruuid", "errorcode", "errornumber", "qaerrorid"
    ) if key in headers), None)
    category_key = next((headers[key] for key in (
        "errorcategory", "category", "errorreason", "reason", "errordescription", "description", "remarks"
    ) if key in headers), None)

    if not all((date_key, username_key, name_key, uuid_key, category_key)):
        raise ValueError(
            "Required columns not found. Include Date, Emp ID, Agent Name, UUID, and error_category."
        )

    parsed = []
    for row_number, row in enumerate(raw_rows, start=2):
        values = [row.get(key) for key in (date_key, username_key, name_key, uuid_key, category_key)]
        if not any(str(value or "").strip() for value in values):
            continue
        error_date = parse_error_date(row.get(date_key))
        username = str(row.get(username_key) or "").strip()
        agent_name = str(row.get(name_key) or "").strip()
        error_id = str(row.get(uuid_key) or "").strip()
        category = str(row.get(category_key) or "").strip()
        if not all((error_date, username, agent_name, error_id, category)):
            raise ValueError(
                f"Row {row_number} is missing Date, Emp ID, Agent Name, UUID, or error_category."
            )
        if error_date.strftime("%Y-%m") != month:
            raise ValueError(
                f"Row {row_number} has date {error_date.isoformat()}, which is outside the selected month {month}."
            )
        parsed.append({
            "error_date": error_date,
            "username": username,
            "agent_name": agent_name,
            "error_id": error_id,
            "category": category
        })
    if not parsed:
        raise ValueError("No complete error records were found.")
    return parsed


def fetch_google_sheet(sheet_url):
    export_url = google_sheet_csv_url(sheet_url)
    req = Request(export_url, headers={"User-Agent": "Mozilla/5.0"})
    try:
        with urlopen(req, timeout=25) as response:
            content = response.read(10 * 1024 * 1024 + 1)
    except Exception as exc:
        raise ValueError(
            "Could not read the Google Sheet. Set sharing to Anyone with the link - Viewer, then try again."
        ) from exc
    if len(content) > 10 * 1024 * 1024:
        raise ValueError("The Google Sheet is larger than the 10 MB import limit.")
    return content


def fetch_sharepoint_file(file_url):
    """Fetch SharePoint files only when the sharing link permits server-side access."""
    parsed = urlparse(file_url)
    host = (parsed.hostname or "").lower()
    if parsed.scheme != "https" or not (host == "sharepoint.com" or host.endswith(".sharepoint.com") or host == "1drv.ms" or host.endswith(".1drv.ms")):
        raise ValueError("Enter a valid HTTPS SharePoint sharing link.")
    download_url = file_url
    separator = "&" if "?" in download_url else "?"
    if "download=" not in download_url.lower():
        download_url += separator + "download=1"
    req = Request(download_url, headers={"User-Agent": "Mozilla/5.0", "Accept": "*/*"})
    try:
        with urlopen(req, timeout=30) as response:
            content = response.read(10 * 1024 * 1024 + 1)
            content_type = response.headers.get("Content-Type", "").lower()
    except Exception as exc:
        raise ValueError("SharePoint could not provide this file. Private company files require Microsoft 365 sign-in integration; upload the file from your device or use a permitted sharing link.") from exc
    if len(content) > 10 * 1024 * 1024:
        raise ValueError("The SharePoint file is larger than the 10 MB import limit.")
    if "text/html" in content_type or content.lstrip().lower().startswith((b"<!doctype html", b"<html")):
        raise ValueError("SharePoint returned a sign-in page. This private file needs Microsoft 365 sign-in integration, or upload it from your device.")
    if content.startswith(bytes((80, 75, 3, 4))):
        filename = "sharepoint.xlsx"
    elif content.startswith(bytes((208, 207, 17, 224))):
        filename = "sharepoint.xls"
    else:
        filename = "sharepoint.csv"
    return content, filename


def import_agent_errors(tl_id, month, parsed_rows, source):
    agents = User.query.filter_by(role="agent", tl_id=tl_id).all()
    agents_by_username = {agent.username.strip().casefold(): agent for agent in agents}
    imported = 0
    skipped_usernames = set()
    seen = set()

    # Replace this source's previous import only after the new file has parsed.
    AgentError.query.filter_by(source_id=source.id).delete(synchronize_session=False)
    for row in parsed_rows:
        agent = agents_by_username.get(row["username"].casefold())
        if not agent:
            skipped_usernames.add(row["username"])
            continue
        key = (agent.id, row["error_id"].casefold())
        if key in seen:
            continue
        seen.add(key)
        db.session.add(AgentError(
            source_id=source.id,
            tl_id=tl_id,
            agent_id=agent.id,
            month=month,
            error_id=row["error_id"][:100],
            reason=row["category"],
            error_date=row["error_date"],
            agent_name=row["agent_name"][:150],
            category=row["category"]
        ))
        imported += 1

    if not imported:
        db.session.rollback()
        raise ValueError("No rows matched active or assigned agents. Check the Emp ID values against agent usernames.")
    db.session.commit()
    return imported, sorted(skipped_usernames)


def save_error_spreadsheet_import(tl_id, month, uploaded_file=None, spreadsheet_url="", existing=None):
    datetime.strptime(month, "%Y-%m")
    if uploaded_file and uploaded_file.filename:
        filename = uploaded_file.filename
        if len(filename) > 255:
            raise ValueError("The uploaded filename is too long.")
        file_bytes = uploaded_file.read(10 * 1024 * 1024 + 1)
        if len(file_bytes) > 10 * 1024 * 1024:
            raise ValueError("The file is larger than the 10 MB import limit.")
        parsed_rows = read_error_rows(file_bytes, filename, month)
        source_url = ""
        source_name = filename
    elif spreadsheet_url:
        if len(spreadsheet_url) > 2000:
            raise ValueError("The spreadsheet link is too long.")
        if "docs.google.com/spreadsheets/" in spreadsheet_url:
            file_bytes = fetch_google_sheet(spreadsheet_url)
            filename = "sheet.csv"
            source_name = "Google Sheets"
        else:
            file_bytes, filename = fetch_sharepoint_file(spreadsheet_url)
            source_name = "SharePoint file"
        parsed_rows = read_error_rows(file_bytes, filename, month)
        source_url = spreadsheet_url
    else:
        raise ValueError("Choose a file from your device or enter a Google Sheets/SharePoint spreadsheet link.")

    item = existing or ErrorSpreadsheet.query.filter_by(tl_id=tl_id, month=month).first()
    if item is None:
        item = ErrorSpreadsheet(tl_id=tl_id, month=month, spreadsheet_url=source_url,
                                source_name=source_name)
        db.session.add(item)
        db.session.flush()
    else:
        item.month = month
        item.spreadsheet_url = source_url
        item.source_name = source_name

    imported, skipped = import_agent_errors(tl_id, month, parsed_rows, item)
    return item, imported, skipped


@app.route("/tl/error-spreadsheets", methods=["GET", "POST"])
@login_required
@role_required("tl")
def tl_error_spreadsheets():
    if request.method == "POST":
        try:
            month = request.form["month"].strip()
            item, imported, skipped = save_error_spreadsheet_import(
                current_user.id, month,
                uploaded_file=request.files.get("spreadsheet_file"),
                spreadsheet_url=request.form.get("spreadsheet_url", "").strip()
            )
            flash(f"Imported {imported} error records for {month}.", "success")
            if skipped:
                flash("Rows skipped because these usernames are not assigned to your team: " + ", ".join(skipped[:10]), "warning")
        except (KeyError, ValueError) as exc:
            db.session.rollback()
            flash(str(exc) or "Please provide a valid month and spreadsheet.", "danger")
        except Exception:
            db.session.rollback()
            app.logger.exception("Error importing TL spreadsheet")
            flash("The spreadsheet could not be imported. Check its format and try again.", "danger")
        return redirect(url_for("tl_error_spreadsheets"))

    edit_id = request.args.get("edit", type=int)
    editing = None
    if edit_id:
        editing = ErrorSpreadsheet.query.filter_by(
            id=edit_id, tl_id=current_user.id
        ).first()
        if not editing:
            flash("That spreadsheet was not found in your team.", "danger")
            return redirect(url_for("tl_error_spreadsheets"))
    spreadsheets = ErrorSpreadsheet.query.filter_by(
        tl_id=current_user.id
    ).order_by(ErrorSpreadsheet.month.desc()).all()
    return render_template("tl_error_spreadsheets.html", spreadsheets=spreadsheets, editing=editing)


@app.route("/tl/error-spreadsheets/<int:spreadsheet_id>/edit", methods=["POST"])
@login_required
@role_required("tl")
def edit_error_spreadsheet(spreadsheet_id):
    item = ErrorSpreadsheet.query.filter_by(id=spreadsheet_id, tl_id=current_user.id).first()
    if not item:
        flash("That spreadsheet was not found in your team.", "danger")
        return redirect(url_for("tl_error_spreadsheets"))
    try:
        month = request.form["month"].strip()
        duplicate = ErrorSpreadsheet.query.filter(
            ErrorSpreadsheet.tl_id == current_user.id,
            ErrorSpreadsheet.month == month,
            ErrorSpreadsheet.id != item.id
        ).first()
        if duplicate:
            raise ValueError("A spreadsheet already exists for that month.")
        _, imported, skipped = save_error_spreadsheet_import(
            current_user.id, month,
            uploaded_file=request.files.get("spreadsheet_file"),
            spreadsheet_url=request.form.get("spreadsheet_url", "").strip(),
            existing=item
        )
        flash(f"Updated {month}: imported {imported} error records.", "success")
        if skipped:
            flash("Rows skipped because these usernames are not assigned to your team: " + ", ".join(skipped[:10]), "warning")
    except (KeyError, ValueError) as exc:
        db.session.rollback()
        flash(str(exc) or "Please provide a valid month and spreadsheet.", "danger")
        return redirect(url_for("tl_error_spreadsheets", edit=item.id))
    except Exception:
        db.session.rollback()
        app.logger.exception("Error updating TL spreadsheet")
        flash("The spreadsheet could not be imported. Check its format and try again.", "danger")
        return redirect(url_for("tl_error_spreadsheets", edit=item.id))
    return redirect(url_for("tl_error_spreadsheets"))


@app.route("/tl/error-spreadsheets/<int:spreadsheet_id>/delete", methods=["POST"])
@login_required
@role_required("tl")
def delete_error_spreadsheet(spreadsheet_id):
    item = ErrorSpreadsheet.query.filter_by(id=spreadsheet_id, tl_id=current_user.id).first()
    if not item:
        flash("That spreadsheet was not found in your team.", "danger")
    else:
        db.session.delete(item)
        db.session.commit()
        flash("Spreadsheet and its imported error records deleted.", "success")
    return redirect(url_for("tl_error_spreadsheets"))


@app.route("/agent/error-management")
@login_required
@role_required("agent")
def agent_error_management():
    records = AgentError.query.filter_by(agent_id=current_user.id).order_by(
        AgentError.month.desc(), AgentError.error_date.desc(), AgentError.error_id.asc()
    ).all()
    grouped = {}
    for record in records:
        grouped.setdefault(record.month, []).append(record)
    errors_by_month = [
        (month, datetime.strptime(month, "%Y-%m").strftime("%B %Y"), rows)
        for month, rows in grouped.items()
    ]
    return render_template("agent_error_management.html", errors_by_month=errors_by_month)


# -------------------------------- TEAM LEAD ---------------------------------
@app.route("/tl/team")
@login_required
@role_required("tl")
def tl_team():
    period, start, end = selected_period()
    agents = my_agents()
    rows, team_calls, team_approvals, team_errors, team_workdays = [], 0, 0, 0, 0
    daily = {}
    for agent in agents:
        k = agent_kpis(agent.id, start, end)
        rows.append({"agent": agent, "k": k})
        team_calls += k["total_calls"]
        team_approvals += k["total_appr"]
        team_errors += k["total_err"]
        team_workdays += k["working_days"]
        for entry in k["entries"]:
            bucket = daily.setdefault(entry.entry_date, [0, 0, 0])
            bucket[0] += entry.total_calls
            bucket[1] += entry.approvals
            bucket[2] += entry.errors
    days = sorted(daily)
    chart = {
        "labels": [d.strftime("%d %b") for d in days],
        "calls": [daily[d][0] for d in days],
        "approval_pct": [round(daily[d][1] * 100 / daily[d][0], 2) if daily[d][0] else 0 for d in days],
        "error_pct": [round(daily[d][2] * 100 / daily[d][1], 2) if daily[d][1] else 0 for d in days],
        "agent_names": [r["agent"].full_name for r in rows],
        "agent_calls": [r["k"]["total_calls"] for r in rows],
    }
    team = {
        "total_calls": team_calls, "total_approvals": team_approvals,
        "total_errors": team_errors,
        "avg_calls": round(team_calls / team_workdays, 1) if team_workdays else 0,
        "avg_approvals": round(team_approvals / team_workdays, 1) if team_workdays else 0,
        "appr_pct": round(team_approvals * 100 / team_calls, 2) if team_calls else 0,
        "err_pct": round(team_errors * 100 / team_approvals, 2) if team_approvals else 0,
        "agent_count": len(agents), "working_days": team_workdays,
        "target": get_target(start.strftime("%Y-%m"))
    }
    return render_template("tl_dashboard.html", mode="team", team=team, rows=rows,
                           chart=chart, agents=agents, y=start.year, m=start.month,
                           period=period, start=start, end=end,
                           month_name=cal.month_name[start.month])


@app.route("/tl/agent/<int:agent_id>")
@login_required
@role_required("tl")
def tl_agent(agent_id):
    agent = db.session.get(User, agent_id)
    if not agent or agent.role != "agent" or agent.tl_id != current_user.id:
        return "Access denied", 403
    period, start, end = selected_period()
    k = agent_kpis(agent.id, start, end)
    return render_template("tl_dashboard.html", mode="agent", agent=agent, k=k,
                           agents=my_agents(), y=start.year, m=start.month,
                           period=period, start=start, end=end,
                           month_name=cal.month_name[start.month], team=None, rows=None,
                           chart=k["chart"])


@app.route("/tl/agent/<int:agent_id>/entry", methods=["POST"])
@login_required
@role_required("tl")
def tl_update_agent_entry(agent_id):
    agent = db.session.get(User, agent_id)
    if not agent or agent.role != "agent" or agent.tl_id != current_user.id:
        return "Access denied", 403
    try:
        entry_date = datetime.strptime(request.form["entry_date"], "%Y-%m-%d").date()
        calls, approvals, errors = (int(request.form[x]) for x in ("total_calls", "approvals", "errors"))
        if entry_date > date.today() or min(calls, approvals, errors) < 0 or approvals > calls or errors > calls:
            raise ValueError
        entry = DailyEntry.query.filter_by(agent_id=agent.id, entry_date=entry_date).first()
        if entry:
            old = f"{entry.total_calls}/{entry.approvals}/{entry.errors}"
            entry.total_calls, entry.approvals, entry.errors = calls, approvals, errors
        else:
            old = "new entry"
            entry = DailyEntry(agent_id=agent.id, entry_date=entry_date,
                               total_calls=calls, approvals=approvals, errors=errors)
            db.session.add(entry)
            db.session.flush()
        db.session.add(EntryHistory(entry_id=entry.id, changed_by=current_user.id,
                                    old_values=old, new_values=f"{calls}/{approvals}/{errors}"))
        db.session.commit()
        flash("Agent performance updated.", "success")
    except (KeyError, ValueError):
        db.session.rollback()
        flash("Invalid entry. Approvals and errors cannot exceed calls.", "danger")
    return redirect(url_for("tl_agent", agent_id=agent.id, period="day",
                            date=request.form.get("entry_date", date.today().isoformat())))


@app.route("/tl/history/<int:agent_id>")
@login_required
@role_required("tl", "admin")
def entry_history(agent_id):
    agent = db.session.get(User, agent_id)
    if not agent:
        return jsonify({"error": "Agent not found"}), 404
    if current_user.role == "tl" and agent.tl_id != current_user.id:
        return jsonify({"error": "Access denied"}), 403
    entry_ids = [e.id for e in DailyEntry.query.filter_by(agent_id=agent_id).all()]
    history = EntryHistory.query.filter(EntryHistory.entry_id.in_(entry_ids)).order_by(
        EntryHistory.changed_at.desc()).limit(200).all() if entry_ids else []
    return jsonify([{
        "when": h.changed_at.strftime("%d %b %Y %H:%M"),
        "by": h.editor.full_name if h.editor else "Unknown",
        "old": h.old_values, "new": h.new_values
    } for h in history])


# ----------------------------- EXCEL EXPORT --------------------------------
@app.route("/export/team.xlsx")
@login_required
@role_required("tl", "admin")
def export_team():
    year, month = parse_month()
    if current_user.role == "tl":
        agents = my_agents()
    else:
        agents = User.query.filter_by(role="agent", active=True).order_by(User.full_name).all()
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Daily Performance"
    sheet.append(["Agent", "Date", "Total Calls", "Approvals", "Errors", "Approval %", "Error %"])
    for agent in agents:
        first, last = month_bounds(year, month)
        entries = DailyEntry.query.filter(
            DailyEntry.agent_id == agent.id,
            DailyEntry.entry_date >= first, DailyEntry.entry_date <= last
        ).order_by(DailyEntry.entry_date).all()
        for e in entries:
            sheet.append([agent.full_name, e.entry_date.isoformat(), e.total_calls,
                          e.approvals, e.errors, e.approval_pct, e.error_pct])
    output = BytesIO()
    workbook.save(output)
    output.seek(0)
    return send_file(output, as_attachment=True,
                     download_name=f"performance_{year}_{month:02d}.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.context_processor
def inject_now_label():
    return {"now_label": date.today().strftime("%A, %d %B %Y")}


# ------------------------------- STARTUP -----------------------------------
def initialize_database():
    db.create_all()

    user_columns = {column["name"] for column in inspect(db.engine).get_columns("user")}
    with db.engine.begin() as connection:
        if "email" not in user_columns:
            connection.exec_driver_sql('ALTER TABLE "user" ADD COLUMN email VARCHAR(254)')
        if "email_verified" not in user_columns:
            connection.exec_driver_sql('ALTER TABLE "user" ADD COLUMN email_verified BOOLEAN NOT NULL DEFAULT FALSE')
        connection.exec_driver_sql('CREATE UNIQUE INDEX IF NOT EXISTS uq_user_verified_email ON "user"(email)')

    # Add the roster type to existing deployments without losing saved dates.
    roster_columns = {column["name"] for column in inspect(db.engine).get_columns("roster")}
    if "roster_type" not in roster_columns:
        with db.engine.begin() as connection:
            connection.exec_driver_sql(
                "ALTER TABLE roster ADD COLUMN roster_type VARCHAR(20) NOT NULL DEFAULT 'weekoff'"
            )
    if "agent_entered" not in roster_columns:
        with db.engine.begin() as connection:
            connection.exec_driver_sql(
                "ALTER TABLE roster ADD COLUMN agent_entered BOOLEAN NOT NULL DEFAULT FALSE"
            )

    spreadsheet_columns = {column["name"] for column in inspect(db.engine).get_columns("error_spreadsheet")}
    if "source_name" not in spreadsheet_columns:
        with db.engine.begin() as connection:
            connection.exec_driver_sql("ALTER TABLE error_spreadsheet ADD COLUMN source_name VARCHAR(255)")

    error_columns = {column["name"] for column in inspect(db.engine).get_columns("agent_error")}
    with db.engine.begin() as connection:
        if "error_date" not in error_columns:
            connection.exec_driver_sql("ALTER TABLE agent_error ADD COLUMN error_date DATE")
        if "agent_name" not in error_columns:
            connection.exec_driver_sql("ALTER TABLE agent_error ADD COLUMN agent_name VARCHAR(150)")
        if "category" not in error_columns:
            connection.exec_driver_sql("ALTER TABLE agent_error ADD COLUMN category TEXT")

    if not User.query.filter_by(username="admin").first():
        admin = User(
            username="admin",
            full_name="Administrator",
            role="admin"
        )
        admin.set_password(
            os.environ.get("INITIAL_ADMIN_PASSWORD", "admin123")
        )
        db.session.add(admin)
        db.session.commit()


# Initialize database for both local and Render deployments
with app.app_context():
    initialize_database()


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=True)
