The Excel Datagrid is a supporting project to GL21 Open Source project. Its objective is to provide a good datagrid in Excel which can edit the Tags table.

Requirements
To run this project you must have access to an instance of SQL Server 2005. The xls file is made in and english version of Excel from Office 2003.

How to run the project:
Download the project files to a directory. From a command line window in the same directory type 'sqlcmd -i setupTesttag.sql'.
Run Excel
Open the testTagDb.xls spreadsheet. If you have the macro security setting to medium, Excel should ask you to use the macros. Click OK or Yes. If you don't get question you may have to adjust Excel Macro settings accordingly.
Open the VBA editor by pressing Alt+F11
Search for and change the 'WSID=YOUR\_SERVERNAME' to reflect your machines name.

Please use the discussion group if you have problems or suggestions.