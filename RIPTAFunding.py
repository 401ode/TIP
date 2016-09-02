import pypyodbc
import pandas as pd
import numpy as np

pypyodbc.lowercase = False
conn = pypyodbc.connect(
    r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};" +
    r"Dbq=C:\Users\Nicholas.Tomlin\Documents\GitHub\TIP\TIP17-25.accdb;")

munis = pd.read_sql("SELECT TIPID, Municipalities FROM IDmunicipalities", conn);
df1 = pd.read_sql("SELECT ID, TIPID, FY17state, FY175307, FY175339, FY175310, FY175311, FY17CMAQ, FY175337, FY18state, FY185307, FY185339, FY185310, FY185311, FY18CMAQ, FY185337, FY19state, FY195307, FY195339, FY195310, FY195311, FY19CMAQ, FY195337, FY20state, FY205307, FY205339, FY205310, FY205311, FY20CMAQ, FY205337, FY21state, FY215307, FY215339, FY215310, FY215311, FY21CMAQ, FY215337, FY22state, FY225307, FY225339, FY225310, FY225311, FY22CMAQ, FY23state, FY235307, FY235339, FY235310, FY235311, FY23CMAQ, FY24state, FY245307, FY245339, FY245310, FY245311, FY24CMAQ, FY25state, FY255307, FY255339, FY255310, FY255311, FY25CMAQ, TIPprogram FROM RIPTAFunding", conn);
df2 = pd.read_sql("SELECT ID, TIPID, FY16state, FY165307, FY165339, FY165310, FY165311, FY16CMAQ, FY16TigerGrant, FY16RICAPhwm, FY165337 FROM FY16_RIPTAFunding", conn);

df = pd.merge(df1, df2, on='TIPID', how='outer');
df_final = pd.merge(munis, df, how='right', on='TIPID');
df_final.to_csv(r'C:\Users\Nicholas.Tomlin\Documents\GitHub\TIP\RIPTA Funding\RIPTA.csv', index=False);

conn.close()
