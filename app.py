from flask import Flask, render_template, request, redirect, url_for, send_file, abort, jsonify
from flask_sqlalchemy import SQLAlchemy
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from datetime import datetime, timezone
import os
import re
from werkzeug.utils import secure_filename
import logging
from dateutil.parser import parse  # Requires: pip install python-dateutil

# Initialize Flask app
app = Flask(__name__)

# Configuration for PythonAnywhere (replace 'mytypist' with your actual username)
BASE_DIR = os.path.abspath(os.path.dirname(__file__))

# BASE_DIR = '/home/mytypist/mytypist_app'
app.config['SQLALCHEMY_DATABASE_URI'] = f'sqlite:///{os.path.join(BASE_DIR, "db", "db.sqlite")}'
app.config['UPLOAD_FOLDER'] = os.path.join(BASE_DIR, 'uploads')
app.config['GENERATED_FOLDER'] = os.path.join(BASE_DIR, 'generated')
app.config['ADMIN_KEY'] = os.environ.get('ADMIN_KEY', 'secretkey123')  # Set this in PythonAnywhere Web tab
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

# Set up logging
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

# Initialize database
db = SQLAlchemy(app)

# Ensure directories exist
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['GENERATED_FOLDER'], exist_ok=True)
os.makedirs(os.path.join(BASE_DIR, 'db'), exist_ok=True)

# **Database Models**
class Template(db.Model):
    __tablename__ = 'template'
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False)
    type = db.Column(db.String(50), nullable=False)
    file_path = db.Column(db.String(200), nullable=False)
    font_family = db.Column(db.String(50), nullable=False)
    font_size = db.Column(db.Integer, nullable=False)
    is_active = db.Column(db.Boolean, default=True)
    placeholders = db.relationship('Placeholder', back_populates='template', cascade="all, delete-orphan")
    created_documents = db.relationship('CreatedDocument', back_populates='template', cascade="all, delete-orphan")

class Placeholder(db.Model):
    __tablename__ = 'placeholder'
    id = db.Column(db.Integer, primary_key=True)
    template_id = db.Column(db.Integer, db.ForeignKey('template.id'), nullable=False)
    name = db.Column(db.String(50), nullable=False)
    paragraph_index = db.Column(db.Integer, nullable=False)
    start_run_index = db.Column(db.Integer, nullable=False)
    end_run_index = db.Column(db.Integer, nullable=False)
    bold = db.Column(db.Boolean, default=False)
    italic = db.Column(db.Boolean, default=False)
    underline = db.Column(db.Boolean, default=False)
    casing = db.Column(db.String(20), default="none")
    template = db.relationship('Template', back_populates='placeholders')

class CreatedDocument(db.Model):
    __tablename__ = 'created_document'
    id = db.Column(db.Integer, primary_key=True)
    template_id = db.Column(db.Integer, db.ForeignKey('template.id'), nullable=False)
    user_name = db.Column(db.String(100), nullable=False)
    file_path = db.Column(db.String(200), nullable=False)
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))
    template = db.relationship('Template', back_populates='created_documents')

# **Helper Functions**
def ordinal(n):
    """Convert a number to its ordinal form (e.g., 1 -> 1st, 2 -> 2nd)."""
    if 11 <= (n % 100) <= 13:
        suffix = 'th'
    else:
        suffix = ['th', 'st', 'nd', 'rd', 'th'][min(n % 10, 4)]
    return str(n) + suffix

def format_date(date_string, template_type):
    """Format a date string based on the template type."""
    try:
        date_obj = parse(date_string)
        day = ordinal(date_obj.day)
        month = date_obj.strftime("%B")
        year = date_obj.year
        if template_type == "letter":
            return f"{day} {month}, {year}"
        elif template_type == "affidavit":
            return f"{day} of {month}, {year}"
        return f"{date_obj.day} {month} {year}"
    except ValueError:
        logger.warning(f"Invalid date format: {date_string}")
        return date_string

def extract_placeholders(doc):
    """Extract placeholders like ${name} from a Word document."""
    placeholders = []
    placeholder_pattern = re.compile(r'\$\{([^}]+)\}')
    for p_idx, paragraph in enumerate(doc.paragraphs):
        full_text = ''.join(run.text for run in paragraph.runs)
        matches = placeholder_pattern.finditer(full_text)
        for match in matches:
            placeholder_name = match.group(1)
            start_pos = match.start()
            end_pos = match.end()
            current_pos = 0
            start_run_idx = end_run_idx = None
            bold = italic = underline = False
            for r_idx, run in enumerate(paragraph.runs):
                run_start = current_pos
                run_end = current_pos + len(run.text)
                if start_run_idx is None and run_start <= start_pos < run_end:
                    start_run_idx = r_idx
                    bold = run.font.bold or False
                    italic = run.font.italic or False
                    underline = run.font.underline or False
                if run_start < end_pos <= run_end:
                    end_run_idx = r_idx
                    break
                current_pos = run_end
            if start_run_idx is not None and end_run_idx is not None:
                placeholders.append({
                    'paragraph_index': p_idx,
                    'start_run_index': start_run_idx,
                    'end_run_index': end_run_idx,
                    'name': placeholder_name,
                    'bold': bold,
                    'italic': italic,
                    'underline': underline,
                    'casing': 'none'
                })
    return placeholders

def detect_document_font(doc):
    """Detect the most common font and size in a document."""
    font_counts = {}
    for para in doc.paragraphs:
        for run in para.runs:
            if run.font.name and run.font.size:
                key = (run.font.name, int(run.font.size.pt))
                font_counts[key] = font_counts.get(key, 0) + 1
    if font_counts:
        return max(font_counts.items(), key=lambda x: x[1])[0]
    return "Times New Roman", 12

def allowed_file(filename):
    """Check if a file has a .docx extension."""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() == 'docx'

def set_default_font(doc, font_name, font_size):
    """Set the default font for a document."""
    style = doc.styles['Normal']
    font = style.font
    font.name = font_name
    font.size = Pt(font_size)

def remove_empty_runs(doc):
    """Remove empty runs from a document to clean up formatting."""
    for para in doc.paragraphs:
        p = para._element
        runs = list(p.findall('.//w:r', namespaces=p.nsmap))
        for run in runs:
            t = run.find('.//w:t', namespaces=run.nsmap)
            if t is not None and t.text == '':
                p.remove(run)

# **Routes**
@app.route('/')
def index():
    """Display the homepage with template types and recent documents."""
    types = db.session.query(Template.type).distinct().filter(Template.is_active == True).all()
    types = [t[0] for t in types]
    page = int(request.args.get('page', 1))
    per_page = 10
    recent_docs = CreatedDocument.query.filter(CreatedDocument.template.has(is_active=True))\
        .order_by(CreatedDocument.created_at.desc())\
        .offset((page-1)*per_page).limit(per_page).all()
    total_docs = CreatedDocument.query.filter(CreatedDocument.template.has(is_active=True)).count()
    total_pages = (total_docs + per_page - 1) // per_page
    return render_template('index.html', types=types, recent_docs=recent_docs,
                         page=page, total_pages=total_pages, admin_key=app.config['ADMIN_KEY'])

@app.route('/templates')
def get_templates():
    """Return a JSON list of templates for a given type."""
    type_ = request.args.get('type')
    if not type_:
        return jsonify({'error': 'Type is required'}), 400
    templates = Template.query.filter_by(type=type_, is_active=True).all()
    return jsonify([{'id': t.id, 'name': t.name} for t in templates])

@app.route('/create/<int:template_id>')
def create(template_id):
    """Render the document creation page for a specific template."""
    template = Template.query.filter_by(id=template_id, is_active=True).first_or_404()
    placeholders = Placeholder.query.filter_by(template_id=template_id)\
        .order_by(Placeholder.paragraph_index, Placeholder.start_run_index).all()
    unique_names = []
    seen = set()
    for ph in placeholders:
        if ph.name not in seen:
            unique_names.append(ph.name)
            seen.add(ph.name)
    return render_template('create.html', template=template, placeholder_names=unique_names)

@app.route('/generate', methods=['POST'])
def generate():
    """Generate a document from a template and user inputs."""
    template_id = request.form['template_id']
    template = Template.query.filter_by(id=template_id, is_active=True).first_or_404()
    user_inputs = {key: request.form[key] for key in request.form if key != 'template_id'}

    doc = Document(os.path.join(app.config['UPLOAD_FOLDER'], template.file_path))
    set_default_font(doc, template.font_family, template.font_size)
    placeholders = Placeholder.query.filter_by(template_id=template.id)\
        .order_by(Placeholder.paragraph_index, Placeholder.start_run_index).all()

    for placeholder in placeholders:
        paragraph = doc.paragraphs[placeholder.paragraph_index]
        if placeholder.start_run_index >= len(paragraph.runs) or placeholder.end_run_index >= len(paragraph.runs):
            logger.warning(f"Invalid run indices for placeholder {placeholder.name} in paragraph {placeholder.paragraph_index}")
            continue

        user_input = user_inputs.get(placeholder.name, "")
        formatted_text = user_input

        if "date" in placeholder.name.lower() or "date_ofbirth" in placeholder.name.lower():
            formatted_text = format_date(user_input, template.type)

        elif "address" in placeholder.name.lower() and template.type == "letter":
            parts = [part.strip() for part in user_input.split(",")]
            if parts:
                if placeholder.start_run_index != placeholder.end_run_index:
                    for r_idx in range(placeholder.start_run_index + 1, placeholder.end_run_index + 1):
                        paragraph.runs[r_idx].text = ""
                run = paragraph.runs[placeholder.start_run_index]
                run.clear()
                for i, part in enumerate(parts):
                    run.add_text(part)
                    if i == len(parts) - 1 or part.endswith("."):
                        if not part.endswith("."):
                            run.add_text(".")
                        break
                    else:
                        run.add_text(",")
                        run.add_break()
                run.font.name = template.font_family
                run.font.size = Pt(template.font_size)
                run.bold = placeholder.bold
                run.italic = placeholder.italic
                run.underline = placeholder.underline
                continue

        else:
            if placeholder.casing == "upper":
                formatted_text = formatted_text.upper()
            elif placeholder.casing == "lower":
                formatted_text = formatted_text.lower()
            elif placeholder.casing == "title":
                formatted_text = formatted_text.title()

        run = paragraph.runs[placeholder.start_run_index]
        if placeholder.start_run_index == placeholder.end_run_index:
            run.text = formatted_text
        else:
            logger.debug(f"Placeholder {placeholder.name} spans multiple runs ({placeholder.start_run_index} to {placeholder.end_run_index})")
            for r_idx in range(placeholder.start_run_index + 1, placeholder.end_run_index + 1):
                paragraph.runs[r_idx].text = ""
            run.text = formatted_text
        run.font.name = template.font_family
        run.font.size = Pt(template.font_size)
        run.bold = placeholder.bold
        run.italic = placeholder.italic
        run.underline = placeholder.underline

    remove_empty_runs(doc)

    user_name = user_inputs.get("name", "Unknown").strip()
    user_name = re.sub(r'\s+', '_', user_name)
    template_name = template.name.strip()
    template_name = re.sub(r'\s+', '_', template_name)
    current_date = datetime.now(timezone.utc).strftime("%Y%m%d")
    file_name = f"{user_name}_{template_name}_{current_date}.docx"
    file_path = os.path.join(app.config['GENERATED_FOLDER'], file_name)
    doc.save(file_path)

    created_doc = CreatedDocument(template_id=template.id, user_name=user_name, file_path=file_name)
    db.session.add(created_doc)
    db.session.commit()
    return send_file(file_path, as_attachment=True)

@app.route('/cdn-cgi/challenge-platform/scripts/jsd/main.js')
def dummy_script():
    """Dummy route to mimic a script request (e.g., for Cloudflare)."""
    return '', 200

@app.route('/download/<int:document_id>')
def download(document_id):
    """Download a previously generated document."""
    doc = CreatedDocument.query.get_or_404(document_id)
    file_path = os.path.join(app.config['GENERATED_FOLDER'], doc.file_path)
    if not os.path.exists(file_path):
        abort(404)
    return send_file(file_path, as_attachment=True)

# **Admin Routes**
@app.route('/admin')
def admin():
    """Display the admin dashboard."""
    key = request.args.get('key')
    if key != app.config['ADMIN_KEY']:
        abort(403)
    templates = Template.query.all()
    total_templates = Template.query.count()
    total_created = CreatedDocument.query.count()
    return render_template('admin.html', templates=templates, total_templates=total_templates,
                         total_created=total_created, admin_key=key)

@app.route('/admin/upload', methods=['POST'])
def upload_template():
    """Upload a new template and extract its placeholders."""
    key = request.form.get('key')
    if key != app.config['ADMIN_KEY']:
        abort(403)
    name = request.form['name']
    type_ = request.form['type']
    file = request.files['file']
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(file_path)
        doc = Document(file_path)
        font_family, font_size = detect_document_font(doc)
        template = Template(name=name, type=type_, file_path=filename,
                          font_family=font_family, font_size=font_size)
        db.session.add(template)
        db.session.commit()
        placeholders = extract_placeholders(doc)
        multi_run_placeholders = [ph for ph in placeholders if ph['start_run_index'] != ph['end_run_index']]
        if multi_run_placeholders:
            names = [ph['name'] for ph in multi_run_placeholders]

        for ph in placeholders:
            placeholder = Placeholder(**ph, template_id=template.id)
            db.session.add(placeholder)
        db.session.commit()
        return redirect(url_for('admin', key=key))
    return "Invalid file", 400

@app.route('/admin/edit/<int:template_id>')
def edit_template(template_id):
    """Render the template edit page."""
    key = request.args.get('key')
    if key != app.config['ADMIN_KEY']:
        abort(403)
    template = Template.query.get_or_404(template_id)
    placeholders = Placeholder.query.filter_by(template_id=template_id).all()
    return render_template('edit.html', template=template, placeholders=placeholders, admin_key=key)

@app.route('/admin/update/<int:template_id>', methods=['POST'])
def update_template(template_id):
    """Update a template's details and placeholder styles."""
    key = request.form.get('key')
    if key != app.config['ADMIN_KEY']:
        abort(403)
    template = Template.query.get_or_404(template_id)
    template.name = request.form['name']
    template.type = request.form['type']
    template.font_family = request.form['font_family']
    template.font_size = int(request.form['font_size'])
    for ph in template.placeholders:
        ph.bold = f'bold_{ph.id}' in request.form
        ph.italic = f'italic_{ph.id}' in request.form
        ph.underline = f'underline_{ph.id}' in request.form
        ph.casing = request.form[f'casing_{ph.id}']
    db.session.commit()
    return redirect(url_for('admin', key=key))

@app.route('/admin/pause/<int:template_id>')
def pause_template(template_id):
    """Pause a template (set is_active to False)."""
    key = request.args.get('key')
    if key != app.config['ADMIN_KEY']:
        abort(403)
    template = Template.query.get_or_404(template_id)
    template.is_active = False
    db.session.commit()
    return redirect(url_for('admin', key=key))

@app.route('/admin/resume/<int:template_id>')
def resume_template(template_id):
    """Resume a paused template (set is_active to True)."""
    key = request.args.get('key')
    if key != app.config['ADMIN_KEY']:
        abort(403)
    template = Template.query.get_or_404(template_id)
    template.is_active = True
    db.session.commit()
    return redirect(url_for('admin', key=key))

@app.route('/delete/<int:document_id>', methods=['GET', 'POST'])
def delete(document_id):
    """Delete a generated document and its file."""
    doc = CreatedDocument.query.get_or_404(document_id)
    file_path = os.path.join(app.config['GENERATED_FOLDER'], doc.file_path)
    if os.path.exists(file_path):
        os.remove(file_path)
    db.session.delete(doc)
    db.session.commit()
    return redirect(url_for('index'))

@app.route('/admin/delete/<int:template_id>')
def delete_template(template_id):
    """Delete a template and its associated data."""
    key = request.args.get('key')
    if key != app.config['ADMIN_KEY']:
        abort(403)
    template = Template.query.get_or_404(template_id)
    db.session.delete(template)
    db.session.commit()
    return redirect(url_for('admin', key=key))

# Run the app locally (not used on PythonAnywhere)
if __name__ == '__main__':
    with app.app_context():
        db.create_all()  # Creates the database tables
    app.run(host='0.0.0.0', port=8000, debug=True)