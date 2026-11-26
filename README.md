# Excel VBA – Outlook Email Automation

## Project Overview

This project automates email distribution using Excel VBA and Microsoft Outlook.

The solution reads recipient information from an Excel worksheet, generates personalized emails with dynamic subject lines, converts department-specific Excel reports into pdf files and sends them to the respective departments heads as attachments, displays summary tables containing overview of the department reports to their respective emails and sends or displays emails based on user-defined settings.

## Business Problem

Sending daily or monthly business reports manually to multiple recipients is time-consuming and repetitive.

This VBA automation solution reduces manual work and improves reporting efficiency.

## Key Features

- Read recipient email addresses from Excel.
- Generate personalized email greetings for department heads.
- Create dynamic email subject lines.
- Generate separate pdf reports for Sales, Finance, HR, and Operations.
- Attach department-specific reports to the corresponding recipients.
- Create professional HTML emails with tables, colors, and bold headings. 
- Send emails automatically using Outlook.
- Display emails before sending for review.
- Prevent duplicate email distribution.
- Validate email addresses.
- Handle runtime errors.
- Maintain an email execution log.
  

## Technologies Used

- Microsoft Excel
- Excel VBA
- Microsoft Outlook Classic
- HTML Email Formatting
- VBA Error Handling
- Excel Worksheets

## Workbook Structure

- **Sales Report** – Contains Sales department data.
- **Finance Report** – Contains Finance department data.
- **HR Report** – Contains HR department data.
- **Operations Report** – Contains Operations department data.
- **Email_List** – Stores department heads' names, email addresses, and active status.
- **Email_Control** – Stores reporting period, sender details, company name, and email sending mode.
- **Email_Log** – Maintains email execution history, status, and error details.

## Email Automation Modes

- **DISPLAY:** Opens individual Outlook email drafts for review before sending.
- **SEND:** Sends department-specific emails automatically through Outlook.

## How to Execute

- Download the Excel workbook.
- Open the workbook in desktop Microsoft Excel.
- Update the month wise department information for each of the department reports. 
- Enable macros if the file is trusted.
- Open the VBA editor using `ALT + F11`.
- Import the `Department_Email_Automation.bas` module.
- Configure department head details in the `Email_List` worksheet.
- Enter the reporting period and email settings in the `Email_Control` worksheet.
- Select `DISPLAY` mode to review the emails.
- Select `SEND` mode to distribute the reports automatically.
- Run the `Run_Department_Email_Automation` macro.
- Review the generated reports and execution history.

## Expected Benefits

- Reduced manual email preparation.
- Faster department-wise report distribution.
- Personalized communication with department heads.
- Improved reporting accuracy.
- Automated attachment generation.
- Prevention of duplicate email distribution.
- Centralized email execution tracking.
- Improved operational efficiency.

## Important Requirements

- Windows desktop Excel.
- Classic Microsoft Outlook desktop for VBA/COM automation.
- Macros enabled.
- Authorized recipient email addresses.
- Department worksheets containing the required report data.

## Disclaimer

- Use fictional sample data when sharing this project publicly.
- Do not include confidential business information or personal email addresses.
- Test the automation in DISPLAY mode before using SEND mode.
- Ensure all recipients are authorized to receive the corresponding department reports.

## Author

Excel VBA Automation Project – Portfolio Demonstration
