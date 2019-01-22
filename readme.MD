Copy the old information from HLAB_Lookup db to the 3 new databases for the 3 MLab environments

- Used during 2017 MLab (Now Allscripts Lab) upgrade to v 16.03 - This is an AutoIT-esque
  method to connect to 2 different databases (in this case, Oracle 11G & MS SQL Server) and
  perform ETL operations.  We copy data from the MS SQL Server to an Oracle 11G
  database.

  The Oracle side are defined as ODBC system DSN's \w Oracle's InstantClient driver.  The SQL
  Server side uses the SQL Server driver + pass-thru auth.
