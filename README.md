Transport Carrier Email Automation
Automation tool built in Python to generate structured payment detail reports and create Outlook draft emails for transport carriers with invoice details and payment confirmations.

This tool helps automate internal financial communication by reducing manual preparation of payment confirmations and ensuring consistent formatting of invoice details.

Overview
The application performs two main tasks:

Generate a cleaned Excel dataset

Reads the operational template (Plantilla.xlsx)
Filters records by the most recent payment date
Extracts relevant invoice and transport data
Generates a clean Excel file ready for communication
Generate Outlook draft emails

Groups records by transport carrier
Builds a formatted HTML table with payment details
Attaches the corresponding PDF payment confirmation
Saves emails automatically in Outlook Drafts
The process significantly reduces repetitive administrative work and improves the reliability of communications with transport providers.

Main Features
Interactive Tkinter interface

Automatic Excel data filtering

Invoice and freight detail formatting

Automatic grouping by transport carrier

Outlook email automation

Automatic PDF attachment detection

Support for multiple workflows:

Anticipo Seco
Saldo Seco
Anticipo Líquidos
Saldo Líquidos

Technologies Used
Python
Pandas
Tkinter
PyWin32 (Outlook integration)
OpenPyXL

Installation
Install the required dependencies:

pip install pandas openpyxl pywin32
Outlook must be installed and configured on the system for the email automation to work.

Running the Application
You can launch the application using the batch file:

lanzar_app.bat
or directly with Python:

python correos_app.py
Application Workflow
Step 1 – Generate the payment detail report
Select:

Mode: ANTICIPO or SALDO
Product type: SECO or LIQUIDOS
Then click:

Generar Planilla
The system will:

Filter the most recent payment date
Generate a clean Excel file with the relevant records
Step 2 – Generate email drafts
Click:

Enviar a Borradores
The application will:

Group records by transport carrier
Generate formatted email content
Attach the corresponding PDF confirmation
Save all emails automatically in Outlook Drafts
Email Content
Each email includes:

Payment confirmation message
HTML table with invoice and freight details
Payment receipt request
Attached payment confirmation PDF
Business Purpose
This project automates a manual operational process used in logistics and finance departments to notify transport carriers about freight payments.

By automating the workflow, the system helps:

Reduce manual administrative tasks
Ensure consistent communication
Prevent errors in invoice reporting
Improve operational efficiency
