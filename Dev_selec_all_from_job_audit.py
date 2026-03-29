import sqlite3

with sqlite3.connect("job_audit.db") as conn:
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    cur.execute("SELECT * FROM audit_log ORDER BY job_id ASC")

    for i,row in enumerate(cur):
        print(i, dict(row),"\n")