# Code from http://stackoverflow.com/questions/25820698/how-do-i-import-an-accdb-file-into-python-and-use-the-data.
# Installed Microsoft Access Database Engine 2010 (linked above).

import pypyodbc
import pandas as pd
import numpy as np

pypyodbc.lowercase = False
conn = pypyodbc.connect(
    r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};" +
    # path to TIP17-25 below should be reconfigured as needed
    r"Dbq=C:\Users\Nicholas.Tomlin\Documents\GitHub\TIP\TIP17-25.accdb;")

# cur = conn.cursor()
# change the SQL query below to alter the data that can be accessed
# cur.execute("SELECT (SELECT * FROM FY17) AS f17, (SELECT * FROM FY18) AS f18, (SELECT * FROM FY19) AS f19, (SELECT * FROM FY20) AS f20");

# while True:
#    row = cur.fetchone()
#    if row is None:
#        break
#    print (u"{2} has ID {0} and TIP ID {1}.".format(row.get("ID"), row.get("TIPID"), row.get("ProjectName")))

df = pd.read_sql("SELECT ID, TIPID, ProjectName FROM AllProjects", conn)
print(df.head());
# cur.close()
conn.close()
