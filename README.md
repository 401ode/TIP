# TIP
A repository for the TIP MS Access database and ETL scripts:

* `TIP17-25.accdb` is the Microsoft Access database with old TIP data
* `Scripts/AllYears.py` creates two files, one for projects from FY17 to FY25, and another specifically for bridge group projects
* `CSVs By Type` contains the two output files from `AllYears.py`
* `Scripts/AnnualReport.py` creates five files, combining FY and BG tables and separating files by year
* `CSVs By Year` contains the five output files from `AnnualReport.py`
* `Scripts/RIPTAFunding.py` returns information from the RIPTAFunding and FY16_RIPTAFunding tables in the MS Access database
* `RIPTA Funding` contains the output of `RIPTAFunding.py`
* `Scripts/ColumnHeaders.py` returns the column headers of a Microsoft Access table, which is good for creating custom SQL queries

Python scripts use the [pypyodbc module](https://pypi.python.org/pypi/pypyodbc) which can be installed with `pip install pypyodbc`. We can connect to the Access database with the following code:

    conn = pypyodbc.connect(
      r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};" +
      r"Dbq=path-to-file-directory\TIP\TIP17-25.accdb;")

Data is read and manipulated using Pandas and the `pandas.read_sql` function. Again, you'll need to install the dependency, this time with `pip install pandas`. Here's an example of the `pandas.read_sql` function, which takes a SQL query and database connection as parameters:

    import pandas as pd
    df17 = pd.read_sql("SELECT * FROM FY17", conn);

More examples and info can be found in `AllYears.py`, `AnnualReport.py`, and `ColumnHeaders.py`.
