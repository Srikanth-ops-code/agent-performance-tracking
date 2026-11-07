import os
import calendar as cal
from datetime import date, datetime
from functools import wraps

from flask import Flask, render_template, request, redirect, url_for, flash, jsonify, send_file
from flask_sqlalchemy import SQLAlchemy
from flask_login import (
    LoginManager, UserMixin, login_user, logout_user, login_required, current_user
)
from werkzeug.security import generate_password_hash, check_password_hash
from sqlalchemy import UniqueConstraint
from io import BytesIO
from openpyxl import Workbook, load_workbook

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

    team_lead = db.relationship("User", remote_side=[id], backref="team_members", uselist=False)

    def set_password(self, password):
        self.password_hash = generate_password_hash(password)

    def check_password(self, password):
        return check_password_hash(self.password_hash, password)


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
    agent = db.relationship("User", backref="roster_days")
    __table_args__ = (UniqueConstraint("agent_id", "off_date", name="uq_agent_off_date"),)


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
    """Weekdays in the selected range, excluding the agent's assigned roster offs."""
    if end_date < start_date:
        return []
    offs = {r.off_date for r in Roster.query.filter(
        Roster.agent_id == agent_id,
        Roster.off_date >= start_date,
        Roster.off_date <= end_date
    ).all()}
    days = []
    current = start_date
    while current <= end_date:
        if current.weekday() < 5 and current not in offs:
            days.append(current)
        current = date.fromordinal(current.toordinal() + 1)
    return days


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
          @app.route("/agent/roster", methods=["GET", "POST"])
@login_required
@role_required("agent")
def agent_roster():
    if request.method == "POST":
        try:
            off_date = datetime.strptime(
                request.form["off_date"], "%Y-%m-%d"
            ).date()

            existing = Roster.query.filter_by(
                agent_id=current_user.id,
                off_date=off_date
            ).first()

            if not existing:
                db.session.add(
                    Roster(
                        agent_id=current_user.id,
                        off_date=off_date
                    )
                )
                db.session.commit()

            flash("Your week-off date has been saved.", "success")

        except (KeyError, ValueError):
            db.session.rollback()
            flash("Choose a valid date.", "danger")

        return redirect(url_for("agent_roster"))

    roster = Roster.query.filter_by(
        agent_id=current_user.id
    ).order_by(Roster.off_date.asc()).all()

    return render_template(
        "agent_roster.html",
        roster=roster,
        today=date.today()
    )
    roster = Roster.query.filter_by(agent_id=current_user.id).order_by(Roster.off_date.asc()).all()
    return render_template("agent_roster.html", roster=roster, today=date.today())


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
    offs = {r.off_date.day for r in Roster.query.filter(
        Roster.agent_id == current_user.id, Roster.off_date >= first, Roster.off_date <= last).all()}
    entry_data = {str(day): {"total_calls": e.total_calls, "approvals": e.approvals, "errors": e.errors}
                  for day, e in entries.items()}
    weeks = cal.Calendar(firstweekday=0).monthdayscalendar(year, month)
    return render_template("calendar.html", y=year, m=month, weeks=weeks, entries=entries,
                           entry_data=entry_data, offs=offs, month_name=cal.month_name[month], today=date.today())


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
