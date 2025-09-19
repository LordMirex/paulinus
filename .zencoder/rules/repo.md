---
description: Repository Information Overview
alwaysApply: true
---

# MyTypist Information

## Summary
MyTypist is a Flask-based web application for generating customized documents from templates. It allows users to upload Word document templates with placeholders, which can then be filled with user-provided information to generate personalized documents. The application includes an admin portal for managing templates and a user interface for creating documents.

## Structure
- **app.py**: Main application file containing all routes and business logic
- **templates/**: HTML templates for the web interface
- **uploads/**: Directory for storing uploaded template files
- **generated/**: Directory for storing generated documents
- **db/**: Contains the SQLite database file
- **venv/**: Python virtual environment

## Language & Runtime
**Language**: Python
**Version**: Python 3.8/3.10 (supports both)
**Framework**: Flask 2.3.3
**Database**: SQLAlchemy 3.0.3 with SQLite
**Template Engine**: Jinja2 (Flask default)

## Dependencies
**Main Dependencies**:
- Flask==2.3.3: Web framework
- Flask-SQLAlchemy==3.0.3: ORM for database operations
- python-docx==0.8.11: Library for working with Word documents
- python-dateutil==2.8.2: Date parsing and formatting
- Werkzeug==2.3.7: WSGI utility library

## Build & Installation
```bash
# Create and activate virtual environment
python -m venv venv
venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Run the application
python app.py
```

## Database
**Type**: SQLite
**Models**:
- Template: Stores document templates with metadata
- Placeholder: Stores placeholder information for templates
- CreatedDocument: Tracks generated documents

## Application Structure
**Main Components**:
- Document template management (upload, edit, activate/deactivate)
- Placeholder detection and management
- Document generation with custom user inputs
- Admin portal for template management

## Frontend
**Framework**: Bootstrap 5.3.0
**Additional Libraries**:
- jQuery 3.6.0
- Animate.css 4.1.1
- Google Fonts (Poppins, Cormorant Garamond)

## Routes
**User Routes**:
- `/`: Homepage with template selection
- `/templates`: API endpoint for template listing
- `/create/<template_id>`: Form for document creation
- `/generate`: Document generation endpoint
- `/download/<document_id>`: Document download endpoint

**Admin Routes**:
- `/admin`: Admin dashboard
- `/admin/upload`: Template upload endpoint
- `/admin/edit/<template_id>`: Template editing interface
- `/admin/update/<template_id>`: Template update endpoint
- `/admin/pause/<template_id>`: Deactivate template
- `/admin/resume/<template_id>`: Activate template
- `/admin/delete/<template_id>`: Delete template

## Deployment
The application is configured for deployment on PythonAnywhere, with appropriate path configurations and environment variable handling.