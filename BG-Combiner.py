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

munis = pd.read_sql("SELECT TIPID, Municipalities FROM IDmunicipalities", conn);

df16 = pd.read_sql("SELECT ID, TIPID, FY16allocation, FY16GasTax, FY16RICAPHIP, FY16RIHMA, FY16RICAPProjects AS RICAPProjects, FY16RICAPFacilities, FY16ProjectCloseouts, FY16IWAYLandSales, FY16GARVEE, FY16TollRevenue, FY16RailwayProgram, FY16HSIP, FY16TAP, FY16NHPP, FY16CMAQ, FY16Planning, FY16STP, FY16NationalFreight, FY16NHTSA, FY16FTA5337 AS FY16FTA, FY16TransitHubBond, FY16TIGER, FY16UnallocatedBonds, FY16other, FY16STBGSA FROM FY16", conn);
bg16 = pd.read_sql("SELECT ID, TIPID, BridgeGroup, FY16allocation, FY16GasTax, FY16RICAPHIP, FY16RIHMA, FY16RICAPprojects AS RICAPProjects, FY16RICAPFacilities, FY16ProjectCloseouts, FY16IWAYLandSales, FY16GARVEE, FY16TollRevenue, FY16RailwayProgram, FY16HSIP, FY16TAP, FY16NHPP, FY16CMAQ, FY16Planning, FY16STP, FY16NationalFreight, FY16NHTSA, FY16FTA, FY16TransitHubBond, FY16TIGERGrant AS FY16TIGER, FY16UnallocatedBonds FROM FY16_BG", conn);
df0 = pd.concat([df16, bg16]);
df = pd.merge(munis, df0, how='outer', on='TIPID');
df.to_csv(r'C:\Users\Nicholas.Tomlin\Documents\GitHub\TIP\CSVs\FY16.csv', index=False);

df17 = pd.read_sql("SELECT * FROM FY17", conn);
bg17 = pd.read_sql("SELECT * FROM FY17_BG", conn);
df1 = pd.concat([df17, bg17]);
df = pd.merge(munis, df1, how='outer', on='TIPID');
df.to_csv(r'C:\Users\Nicholas.Tomlin\Documents\GitHub\TIP\CSVs\FY17.csv', index=False);

df18 = pd.read_sql("SELECT *, FY18Allocation AS Allocation FROM FY18", conn);
bg18 = pd.read_sql("SELECT *, FY18allocation AS Allocation FROM FY18_BG", conn);
df2 = pd.concat([df18, bg18]);
df = pd.merge(munis, df2, how='outer', on='TIPID');
df.to_csv(r'C:\Users\Nicholas.Tomlin\Documents\GitHub\TIP\CSVs\FY18.csv', index=False);

df19 = pd.read_sql("SELECT * FROM FY19", conn);
bg19 = pd.read_sql("SELECT * FROM FY19_BG", conn);
df3 = pd.concat([df19, bg19]);
df = pd.merge(munis, df3, how='outer', on='TIPID');
df.to_csv(r'C:\Users\Nicholas.Tomlin\Documents\GitHub\TIP\CSVs\FY19.csv', index=False);

df20 = pd.read_sql("SELECT * FROM FY20", conn);
bg20 = pd.read_sql("SELECT * FROM FY20_BG", conn);
df4 = pd.concat([df20, bg20]);
df = pd.merge(munis, df4, how='outer', on='TIPID');
df.to_csv(r'C:\Users\Nicholas.Tomlin\Documents\GitHub\TIP\CSVs\FY20.csv', index=False);

df21_25 = pd.read_sql("SELECT * FROM [FY21-25lump]", conn);
bg21_25 = pd.read_sql("SELECT * FROM [FY21-25_BG]", conn);
df5 = pd.concat([df21_25, bg21_25]);
df = pd.merge(munis, df5, how='outer', on='TIPID');
df.to_csv(r'C:\Users\Nicholas.Tomlin\Documents\GitHub\TIP\CSVs\FY21-25.csv', index=False);

# cur.close()
conn.close()
