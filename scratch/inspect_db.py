import sqlite3
import os

db_path = "webapp/phenoldb.sqlite"
if os.path.exists(db_path):
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
    tables = cursor.fetchall()
    print "Tables in database:"
    for t in tables:
        tname = str(t[0])
        print "- " + tname
        cursor.execute("PRAGMA table_info(\"" + tname + "\");")
        cols = [str(c[1]) for c in cursor.fetchall()]
        print "  Columns: " + ", ".join(cols)
    conn.close()
else:
    print "Database not found at " + db_path
