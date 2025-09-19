#!/usr/bin/env python3
"""
Script to clean up invalid templates from the database
"""

import os
import sqlite3
from datetime import datetime

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
DB_PATH = os.path.join(BASE_DIR, 'db', 'db.sqlite')
UPLOADS_DIR = os.path.join(BASE_DIR, 'uploads')

def cleanup_invalid_templates():
    """Remove templates from database if their files don't exist"""
    try:
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        
        # Get all templates
        cursor.execute("SELECT id, name, file_path, is_active FROM template")
        templates = cursor.fetchall()
        
        invalid_template_ids = []
        valid_templates = []
        
        print("Checking template files...")
        for template_id, name, file_path, is_active in templates:
            full_path = os.path.join(UPLOADS_DIR, file_path)
            if os.path.exists(full_path):
                print(f"✓ Valid: {name} -> {file_path}")
                valid_templates.append((template_id, name, file_path))
            else:
                print(f"✗ Missing: {name} -> {file_path}")
                invalid_template_ids.append(template_id)
        
        if invalid_template_ids:
            print(f"\nFound {len(invalid_template_ids)} templates with missing files.")
            
            # Delete placeholders for invalid templates
            placeholders_deleted = 0
            for template_id in invalid_template_ids:
                cursor.execute("SELECT COUNT(*) FROM placeholder WHERE template_id = ?", (template_id,))
                count = cursor.fetchone()[0]
                placeholders_deleted += count
                cursor.execute("DELETE FROM placeholder WHERE template_id = ?", (template_id,))
            
            # Delete invalid templates
            cursor.executemany("DELETE FROM template WHERE id = ?", [(tid,) for tid in invalid_template_ids])
            
            conn.commit()
            print(f"Cleaned up {len(invalid_template_ids)} invalid templates")
            print(f"Cleaned up {placeholders_deleted} associated placeholders")
            
            # Show remaining valid templates
            print(f"\nRemaining valid templates: {len(valid_templates)}")
            for template_id, name, file_path in valid_templates:
                print(f"  - {name}")
        else:
            print("All template files are valid!")
        
        conn.close()
        return True
        
    except Exception as e:
        print(f"Error during cleanup: {str(e)}")
        return False

if __name__ == "__main__":
    success = cleanup_invalid_templates()
    if success:
        print("\nCleanup completed successfully!")
    else:
        exit(1)