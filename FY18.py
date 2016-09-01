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

df17 = pd.read_sql("SELECT * FROM FY17", conn);
bg17 = pd.read_sql("SELECT * FROM FY17_BG", conn);
df1 = pd.concat([df17, bg17]);
df1.to_csv(r'C:\Users\Nicholas.Tomlin\Documents\GitHub\TIP\CSVs\FY17.csv', index=False);

df18 = pd.read_sql("SELECT *, FY18Allocation AS Allocation FROM FY18", conn);
bg18 = pd.read_sql("SELECT *, FY18allocation AS Allocation FROM FY18_BG", conn);
df2 = pd.concat([df18, bg18]);
df2.to_csv(r'C:\Users\Nicholas.Tomlin\Documents\GitHub\TIP\CSVs\FY18.csv', index=False);

df19 = pd.read_sql("SELECT * FROM FY19", conn);
bg19 = pd.read_sql("SELECT * FROM FY19_BG", conn);
df3 = pd.concat([df19, bg19]);
df3.to_csv(r'C:\Users\Nicholas.Tomlin\Documents\GitHub\TIP\CSVs\FY19.csv', index=False);

df20 = pd.read_sql("SELECT * FROM FY20", conn);
bg20 = pd.read_sql("SELECT * FROM FY20_BG", conn);
df4 = pd.concat([df20, bg20]);
df4.to_csv(r'C:\Users\Nicholas.Tomlin\Documents\GitHub\TIP\CSVs\FY20.csv', index=False);

df21_25 = pd.read_sql("SELECT * FROM [FY21-25lump]", conn);
bg21_25 = pd.read_sql("SELECT * FROM [FY21-25_BG]", conn);
df4 = pd.concat([df21_25, bg21_25]);
df4.to_csv(r'C:\Users\Nicholas.Tomlin\Documents\GitHub\TIP\CSVs\FY21-25.csv', index=False);

# cur.close()
conn.close()
