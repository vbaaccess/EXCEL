Projekt: SQLServerImporter
   Project using VBA objects to create the necessary classes.
   
User Story:
   Automating the procedure of downloading data from a database (eg MS SQL Server) to an xls file for further processing / processing.

Project Description:
VBA modules:
  - modMain; main module with main class declaration (clsMain),
  - clsMain; main class; containing main functions and metods,
  - clsConnectionString; responsible for generating connection string data for logging with the use of Active Directory or MS SQL Server authorization,
  - clsSQL; generating SQL queries,
  - clsSettings; reading and writing parameters used in the project,
  - clsHelpr;  
VB User Forms:
  - UserFormSettings; reading and writing project parameters and settings,
  - UserFormMsgBox; user form new wersion of standar MsgBox for project,
SQL:
  - 'DDL - database for testing.sql'; SQL structure for project,
