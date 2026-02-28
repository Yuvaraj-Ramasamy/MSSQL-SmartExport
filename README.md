<p align="center">
  <strong>“Automation is not about replacing people — it’s about empowering them to do more meaningful work.”</strong>
</p>

# MSSQL-SmartExport
Python script to query MSSQL server principals and export results to Excel.

A lightweight Python utility to query SQL Server principals from sys.server_principals and export the results into a clean, formatted Excel report.
Designed for DBAs, developers, and auditors who need a quick way to extract login/user metadata directly from MSSQL.


# 🚀 Features

* Pulls login/user/group details from `sys.server_principals`
* Exports results to Excel (`.xlsx`)
* Auto‑formats the sheet:

    * Thin borders for all cells
    * Auto-adjusted column widths

* Simple, clean Python implementation with pandas, sqlalchemy, and openpyxl
#

# 📂 Repository Structure

```
MSSQL-SmartExport/
│── main.py             # Main script (run this)
│── requirements.txt    # Python dependencies
│── README.md           # Project documentation
│── .gitignore          # Git ignore rules
```
#

# 🧰 Requirements

Install required packages:
```
pip install -r requirements.txt
```

**Dependencies include:**
`pandas`
`sqlalchemy`
`pyodbc`
`openpyxl`

Also ensure you have an appropriate ODBC Driver installed (such as ODBC Driver 17/18 for SQL Server).
#

# ⚙️ Configuration

Inside main.py, update the connection string:
```
Pythonconn_str = (    'mssql+pyodbc://@SERVER,PORT/DATABASE?driver=DRIVER_NAME&trusted_connection=yes')Show more lines
```

**Replace:**

SERVER → SQL Server host
PORT → SQL Server port (1433 default)
DATABASE → Target database
DRIVER_NAME → e.g. ODBC Driver 18 for SQL Server
#

▶️ How to Run
From the project root:
```
python main.py
```

If successful, it will generate:
```
UserQueryResults.xlsx
```

**with:**
Columns: name, type_desc, is_disabled, default_database_name
Borders applied
Auto‑sized columns
#

# 📊 SQL Query Used
```
SQLSELECT
    name,
    type_desc,
    is_disabled,
    default_database_name
FROM sys.server_principals
WHERE type IN ('S', 'U', 'G');Show more lines
```

**This captures:**

S → SQL logins
U → Windows users
G → Windows groups
#

# 🧯 Troubleshooting

**❗ Driver Errors**
Run this to see available drivers:
```
import pyodbc
print(pyodbc.drivers())
```
Ensure your driver= in the connection string matches exactly.

**❗ Empty Rows**
Your SQL login might not have permissions. Ask your DBA or use an elevated account.
#

# 📌 Future Enhancements (Planned)

* CLI arguments for server/db/output file
* Export in CSV/Parquet formats
* Add server role mapping
* Add login permission extraction
#

# 📄 License
MIT License. Feel free to fork, modify, and enhance.
#
