# TIP
A repository for the TIP MS Access database and ETL scripts:

* `TIP17-25.accdb` is the Microsoft Access database with old TIP data
* `access.py` creates two files, one for projects from FY17 to FY25, and another specifically for bridge group projects
* `BG-Combiner` creates five files, combining FY and BG tables and separating files by year
* `CSVs` stores the outputs of `access.py` and `BG-Combiner`

Python scripts use the [pypyodbc module](https://pypi.python.org/pypi/pypyodbc) which can be installed with `pip install pypyodbc`. We can connect to the Access database with the following code:

    conn = pypyodbc.connect(
      r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};" +
      r"Dbq=path-to-file-directory\TIP\TIP17-25.accdb;")

Data is read and manipulated using Pandas and the `pandas.read_sql` function. Again, you'll need to install the dependency, this time with `pip install pandas`. Here's an example of the `pandas.read_sql` function, which takes a SQL query and database connection as parameters:

    import pandas as pd
    df17 = pd.read_sql("SELECT * FROM FY17", conn);

More examples and info can be found in `access.py` and `BG-Combiner.py`.
