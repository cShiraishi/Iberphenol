import sqlite3
import os

db_path = "webapp/phenoldb.sqlite"
if os.path.exists(db_path):
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    cursor.execute("SELECT citation FROM \"references\" LIMIT 10;")
    rows = cursor.fetchall()
    print "Sample Citations:"
    for r in rows:
        print "- " + str(r[0])
    conn.close()
else:
    print "Database not found."
