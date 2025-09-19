#!/usr/bin/env python3
"""
Database migration script to add new fields for batch processing
"""

import sqlite3
import os

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
DB_PATH = os.path.join(BASE_DIR, 'db', 'db.sqlite')

def update_database():
    """Update the database schema with new fields for batch processing"""
    try:
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        
        # Add batch_id column to created_document table if it doesn't exist
        try:
            cursor.execute("ALTER TABLE created_document ADD COLUMN batch_id TEXT")
            print("Added batch_id column to created_document table")
        except sqlite3.OperationalError as e:
            if "duplicate column name" in str(e).lower():
                print("batch_id column already exists in created_document table")
            else:
                raise e
        
        # Create batch_generation table if it doesn't exist
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS batch_generation (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                batch_id TEXT UNIQUE NOT NULL,
                user_name TEXT NOT NULL,
                template_ids TEXT NOT NULL,
                user_inputs TEXT NOT NULL,
                zip_file_path TEXT,
                status TEXT DEFAULT 'pending',
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                completed_at TIMESTAMP
            )
        ''')
        print("Created/verified batch_generation table")
        
        # Add font_name and font_size columns to placeholder table if they don't exist
        try:
            cursor.execute("ALTER TABLE placeholder ADD COLUMN font_name TEXT")
            print("Added font_name column to placeholder table")
        except sqlite3.OperationalError as e:
            if "duplicate column name" in str(e).lower():
                print("font_name column already exists in placeholder table")
            else:
                pass  # Ignore other errors for now
        
        try:
            cursor.execute("ALTER TABLE placeholder ADD COLUMN font_size REAL")
            print("Added font_size column to placeholder table")
        except sqlite3.OperationalError as e:
            if "duplicate column name" in str(e).lower():
                print("font_size column already exists in placeholder table")
            else:
                pass  # Ignore other errors for now
        
        conn.commit()
        conn.close()
        
        print("Database update completed successfully!")
        return True
        
    except Exception as e:
        print(f"Error updating database: {str(e)}")
        return False

if __name__ == "__main__":
    success = update_database()
    if not success:
        exit(1)