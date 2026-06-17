from flask import Flask, render_template, request, session, redirect, url_for,flash
from flask_sqlalchemy import SQLAlchemy
from flask_bcrypt import Bcrypt
from datetime import date,datetime
import os

app = Flask(__name__)

#Config
app.config['SECRET_KEY'] = os.getenv("SESSION_KEY")
app.config["SQLALCHEMY_DATABASE_URI"] = os.getenv("DATABASE_URL")
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False

db = SQLAlchemy(app)
bcrypt = Bcrypt(app)

#DB Layout
class User(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False)
    email = db.Column(db.String(100), unique=True, nullable=False)
    password = db.Column(db.String(150), nullable=False)
    assigned_tasks = db.relationship('Task',foreign_keys='Task.assigned_to',backref='assignee',lazy=True) # Relationships
    created_tasks = db.relationship('Task',foreign_keys='Task.assigned_by',backref='creator',lazy=True) # Relationships

class Task(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    title = db.Column(db.String(850), nullable=False)
    assigned_to = db.Column(db.Integer,db.ForeignKey('user.id'),nullable=False)
    assigned_by = db.Column(db.Integer,db.ForeignKey('user.id'),nullable=False)
    status = db.Column(db.String(20),default="Pending")
    created_date = db.Column(db.Date,default=date.today)
    due_date = db.Column(db.Date,nullable=False)
    completed_date = db.Column(db.Date,nullable=True)

# with app.app_context():
#     db.create_all() #--> this increases cold start latency by 2-5 secs , thus create DB Manually

#Login
@app.route('/', methods=['GET','POST'])
def login():
    if request.method == 'POST':
        email = request.form.get('email')
        password = request.form.get('password')

        try:
            user = User.query.filter_by(email=email).first()

            if user and bcrypt.check_password_hash(user.password,password):
                session["user_id"] = user.id
                return redirect(url_for('dashboard'))
            else:
                flash("Invalid Credentials. Please try again.")
                return render_template('login.html')
        except:
            flash('Internal Application Error')
            return render_template('login.html')
    return render_template('login.html')

# Registration
@app.route('/register', methods=['GET','POST'])
def register():
    if request.method == 'POST':
        try:
            userName = request.form.get('full_name')
            email = request.form.get('email')
            password = bcrypt.generate_password_hash(request.form.get('password')).decode('utf-8')

            new_user = User(name=userName,email=email,password=password)
        
            db.session.add(new_user)
            db.session.commit()
            return redirect(url_for('login'))
        except:
            flash('Please Check information or try again later.')
            return render_template('register.html')
    return render_template('register.html')

#Dashboard
@app.route("/dashboard")
def dashboard():

    if "user_id" not in session:
        return redirect(url_for("login"))

    try:
        currentUser = User.query.get(session["user_id"])

        pending_tasks = Task.query.filter_by(
            assigned_to=currentUser.id,
            status="Pending"
        ).all()

        completed_tasks = Task.query.filter_by(
            assigned_to=currentUser.id,
            status="Completed"
        ).all()

        created_pending = Task.query.filter_by(
            assigned_by=currentUser.id,
            status="Pending"
        ).all()

        created_completed = Task.query.filter_by(
            assigned_by=currentUser.id,
            status="Completed"
        ).all()

        users = User.query.all()

        return render_template(
            "dashboard.html",
            currentUser=currentUser,
            users=users,
            pending=pending_tasks,
            completed=completed_tasks,
            created_pending=created_pending,
            created_completed=created_completed
        )
    except:
        flash('Error loading Data')
        return render_template('dashboard.html')


@app.route('/add_task', methods=['POST'])
def add_task():

    if "user_id" not in session:
        return redirect(url_for("login"))

    title = request.form.get('title')
    assigned_to_name = request.form.get('to')

    assigned_user = User.query.filter_by(
        name=assigned_to_name
    ).first()

    try:
        due_date = datetime.strptime(
            request.form.get("due_date"),
            "%Y-%m-%d"
        ).date()
    except:
        flash('Please enter Target Date')
        return redirect(url_for('dashboard'))

    if title == None or title == "":
        flash('Please enter Task name')
        return redirect(url_for('dashboard'))

    if not assigned_user:
        flash("Please select user")
        return redirect(url_for('dashboard'))

    new_task = Task(
        title=title,
        assigned_to=assigned_user.id,
        assigned_by=session["user_id"],
        due_date=due_date
    )

    db.session.add(new_task)
    db.session.commit()

    flash("Task assigned successfully")

    return redirect(url_for('dashboard'))

@app.route('/complete-task/<int:task_id>', methods=['POST'])
def complete_task(task_id):

    if "user_id" not in session:
        return redirect(url_for("login"))

    task = Task.query.get_or_404(task_id)

    if task.assigned_to != session["user_id"]:
        flash("Unauthorized")
        return redirect(url_for("dashboard"))

    task.status = "Completed"
    task.completed_date = date.today()

    db.session.commit()

    return redirect(url_for('dashboard'))

# @app.route('/delete-task/<int:task_id>', methods=['POST'])
# def delete_task(task_id):

#     if "user_id" not in session:
#         return redirect(url_for("login"))

#     task = Task.query.get_or_404(task_id)

#     if task.assigned_to != session["user_id"]:
#         flash("Unauthorized action.")
#         return redirect(url_for("dashboard"))

#     db.session.delete(task)
#     db.session.commit()

#     flash("Task deleted successfully.")
#     return redirect(url_for("dashboard"))

#Logout
@app.route("/logout")
def logout():
    if "user_id" in session:
        session.clear()
        return redirect(url_for('login'))
    else:
        return redirect(url_for('login'))