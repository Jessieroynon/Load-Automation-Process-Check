# Load Automation Process Check Application for Automation

## Instructions to Download and Run the App

To download the latest version of the Load Automation Process Check Application, please visit the [Releases](https://github.com/Jessieroynon/Load-Automation-Process-Check/releases/tag/v1.3.0) page.

1. Go to the [Releases](https://github.com/Jessieroynon/Load-Automation-Process-Check/releases/tag/v1.3.0) page of this repository.
2. Find the latest release version.
3. Download the app executable file (`.exe`) attached to the release.
4. Run the executable file to install the app on your system.


## ODBC Setup Instructions
To ensure proper functionality of the application, you must set up two ODBC connections: one for the Teradata database and one for the IBM DB2 database.

1. EDWPROD (Teradata)
DSN Name: EDWPROD
Driver: Teradata Database ODBC Driver 17.20

Setup Steps:
a) Open the ODBC Data Source Administrator on your system.
b) Navigate to the "System DSN" tab and click "Add".
c) Select "Teradata Database ODBC Driver 17.20" from the list and click "Finish".
d) In the Data Source Name (DSN) field, enter EDWPROD.
e) Fill in the required connection details, such as server name, username, and password.
f) Click "Test Connection" to ensure the details are correct.
g) Click "OK" to save the DSN.

![image](https://github.com/user-attachments/assets/f4c0e610-773e-40e1-a832-9131ab6ebaa8)

2. ECPROD (IBM DB2)
DSN Name: ECPROD
Driver: IBM DB2 ODBC DRIVER - DB2COPY1

Setup Steps:
a) Open the ODBC Data Source Administrator on your system.
b) Navigate to the "System DSN" tab and click "Add".
c) Select "IBM DB2 ODBC DRIVER - DB2COPY1" from the list and click "Finish".
d) In the Data Source Name (DSN) field, enter ECPROD.
e) Fill in the required connection details, such as server name, database name, username, and password.
f) Click "Test Connection" to ensure the details are correct.
g) Click "OK" to save the DSN.

## Instructions for Using the Load Automation Process Check App
Login: Login to the app using your HMS credentials (G# and Password).
Enter Client Code: Enter a client code into the client code entry box at the top left of the App.
Run Batch Table: Hit the 'Run' button for the Batch Table.
Select Batch Table Row: Select a row from the Batch Table.
Run Stage Table: Hit the 'Run' button for the Stage Table.
Select Stage Table Row: Select a row from the Stage Table.
Load Data: Hit the 'Load Data' button to load the information into the Update Date listboxes.
Update: Hit the 'Update' button at the lower right corner.
The data should be loaded into the LA Process tab of the client's Load Report Update Excel spreadsheet.

## Feedback and Support
If you encounter any issues or have feedback on the application, please feel free to reach out to the developer for assistance.

Contact: Jessie Roynon
Email: jessie.roynon@gainwelltechnologies.com

