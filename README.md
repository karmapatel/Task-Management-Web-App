🗂️ Task Management App

- A simple web-based Task Management system built using Flask and SQLite.
- It allows users to create tasks, assign them to other users, and update task status in a single dashboard interface.

--------------------------------------------------------------------------------------------------------------------

🚀 Features
- User-based task assignment (assign tasks to anyone)
- Create new tasks
- View pending and completed tasks
- Update task status (Pending ↔ Completed)
- Single dashboard for all operations (no multiple pages needed)
- SQLite database integration
- Jinja2 templating for dynamic UI

--------------------------------------------------------------------------------------------------------------------

🛠️ Tech Stack
- Backend: Flask (Python)
- Database: SQLite
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
    - assigned_by
    - assigned_to
    - status

--------------------------------------------------------------------------------------------------------------------