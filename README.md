# MSSQL-SmartExport
Python script to query MSSQL server principals and export results to Excel.

A lightweight Python utility to query SQL Server principals from sys.server_principals and export the results into a clean, formatted Excel report.
Designed for DBAs, developers, and auditors who need a quick way to extract login/user metadata directly from MSSQL.


# 🚀 Features

* Pulls login/user/group details from sys.server_principals
* Exports results to Excel (.xlsx)
* Auto‑formats the sheet:

    * Thin borders for all cells
    * Auto-adjusted column widths

* Simple, clean Python implementation with pandas, sqlalchemy, and openpyxl
