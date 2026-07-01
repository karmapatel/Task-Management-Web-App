🗂️ TaskFlow - Task Management App

- A web-based Task Management system built using Flask and PostgreSQL.
- It allows users to create tasks/projects, assign them to other users, and update task status in dashboard interface.

--------------------------------------------------------------------------------------------------------------------

🚀 Features
- User-based task/project assignment (assign tasks/projects to anyone)
- Create new tasks/projects
- View pending and completed tasks/projects
- Update task status (Pending -> Completed)
- Export your Data in .xlsx format 
- PostgreSQL database integration
- Jinja2 templating for dynamic UI

--------------------------------------------------------------------------------------------------------------------

🛠️ Tech Stack
- Backend: Flask (Python)
- Database: PostgreSQL
- Frontend: HTML, CSS, Jinja2
- ORM: Flask-SQLAlchemy

--------------------------------------------------------------------------------------------------------------------

🗄️ Database

Tables include:

- User
    - id
    - username
    - password

- Task
    - id
    - title
    - assigned_to
    - assigned_by
    - created_date
    - due_date
    - completed_date
    - status

- project
    - id
    - name
    - assigned_to
    - assigned_by

--------------------------------------------------------------------------------------------------------------------