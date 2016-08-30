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

# cur = conn.cursor()
# change the SQL query below to alter the data that can be accessed
# cur.execute("SELECT (SELECT * FROM FY17) AS f17, (SELECT * FROM FY18) AS f18, (SELECT * FROM FY19) AS f19, (SELECT * FROM FY20) AS f20");

# while True:
#    row = cur.fetchone()
#    if row is None:
#        break
#    print (u"{2} has ID {0} and TIP ID {1}.".format(row.get("ID"), row.get("TIPID"), row.get("ProjectName")))
df17 = pd.read_sql("SELECT * FROM FY17", conn);
df18 = pd.read_sql("SELECT * FROM FY18", conn);
df19 = pd.read_sql("SELECT * FROM FY19", conn);
df20 = pd.read_sql("SELECT * FROM FY20", conn);
df21_25 = pd.read_sql("SELECT * FROM [FY21-25lump]", conn);

# reduce() is a left-fold from python 2.7, but it is no longer implemented in 3.5
# so I copied the implementation from the below site:
# https://docs.python.org/2/library/functions.html#reduce
def reduce(function, iterable, initializer=None):
    it = iter(iterable)
    if initializer is None:
        try:
            initializer = next(it)
        except StopIteration:
            raise TypeError('reduce() of empty sequence with no initial value')
    accum_value = initializer
    for x in it:
        accum_value = function(accum_value, x)
    return accum_value

df = reduce(lambda left, right: pd.merge(left, right, how='outer', on='TIPID'), [df17, df18, df19, df20, df21_25])
df = df.drop('ProjectName', 1)
# change output path below as necessary
df.to_csv(r'C:\Users\Nicholas.Tomlin\Documents\GitHub\TIP\CSVs\FY17_25.csv', index=False);
print(df.head());
# cur.close()
conn.close()
