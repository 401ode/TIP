# Code for outputting all of the column headers in a Microsoft Access table.

import pypyodbc
import pandas as pd
import numpy as np

pypyodbc.lowercase = False
conn = pypyodbc.connect(
    r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};" +
    r"Dbq=C:\Users\Nicholas.Tomlin\Documents\GitHub\TIP\TIP17-25.accdb;")

df1 = pd.read_sql("SELECT * FROM RIPTAFunding", conn);
df2 = pd.read_sql("SELECT * FROM FY16_RIPTAFunding", conn);

print(str(list(df1.columns.values)).replace("'",""));
print(str(list(df2.columns.values)).replace("'",""));

conn.close();
