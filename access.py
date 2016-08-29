import pypyodbc
pypyodbc.lowercase = False
conn = pypyodbc.connect(
    r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};" +
    r"Dbq=C:\Users\Nicholas.Tomlin\Desktop\TIP17-25.accdb;")
cur = conn.cursor()
cur.execute("SELECT ID, TIPID, ProjectName FROM AllProjects");
while True:
    row = cur.fetchone()
    if row is None:
        break
    print (u"{2} has ID {0} and TIP ID {1}.".format(row.get("ID"), row.get("TIPID"), row.get("ProjectName")))
cur.close()
conn.close()
