# Reporting_Automation_VBA

## 1. Overview
The Excel application was built to automate the process that the analyst needs to do every day - to import the data of returns, analyse the reason of returns, and make charts to show where the problem is and how to fix it.

The reports contain 2-10 pages with bar/pie charts and tables that are created from pivot tables of different dimentions. The report formats can be modified easily using Excel's pivot table and chart tools.

The UI is showed in "User Interface.pdf" and "Reporting_Automation_PutItOn.xlsm" is the final application.

## 2. UI and Features
The Welcome page shows instructions and three functional buttons:  
__1. Create New Return:__ to input return data manually  
__2. View Return:__ to view the return data  
__3. Generate Report:__ to show a form allowing the user to choose what kind of reports to generate  

The __View Return__ page demostrates a table of return data, including Date, Order ID, Product ID, Vendor, Reasons, etc.

The __Generate Report__ form allows the user the choose three kinds of reports: Overview, Detail, and Vendor Specific report.

After clicking "__Generate Report__" Button, the application will generate a new excel file with a list of charts that is ready to be printed or saved as a PDF file.

The demostrated files contain pseudonymised data to protect the client's infomation.

## 3. Code files
The VBA code is upload in the "__VBA Code__" folder as txt files.  

* "__DataConnection.txt__": Database connections
* "__DataGenerationForTestPurpose.txt__": Create test data
* "__Report.txt__": Generate reports
* "__frmReport.txt__": Code related to the Generate Report form
* "__frmReturnReason.txt__": Code related to the Create New Return form
* "__wsReturnDatabase.txt__": Code related to the Return Database worksheet
* "__wsWelcome.txt__": Code related to the welcome page
