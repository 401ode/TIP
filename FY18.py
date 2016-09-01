# MS Access connection code from
# http://stackoverflow.com/questions/25820698/how-do-i-import-an-accdb-file-into-python-and-use-the-data
# Dependency: Microsoft Access Database Engine 2010 (download link above)

import pypyodbc
import pandas as pd
import numpy as np

pypyodbc.lowercase = False
conn = pypyodbc.connect(
    r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};" +
    # path to TIP17-25 below should be reconfigured as needed
    r"Dbq=C:\Users\Nicholas.Tomlin\Documents\GitHub\TIP\TIP17-25.accdb;")

df18 = pd.read_sql("SELECT *, FY18Allocation AS Allocation FROM FY18", conn);
bg18 = pd.read_sql("SELECT *, FY18allocation AS Allocation FROM FY18_BG", conn)

df = pd.concat([df18, bg18])

# change output path below as necessary
df.to_csv(r'C:\Users\Nicholas.Tomlin\Documents\GitHub\TIP\CSVs\FY18.csv', index=False);

# cur.close()
conn.close()
