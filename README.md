step by step process:

🔁 1. Read Excel Data
Scope: Excel Process Scope

File: employee_data.xlsx

Read Range Activity:

Reads data from Sheet1

Stores it into a DataTable variable named dtStudents

📋 2. For Each Row in Data Table
Loops through each employee's data row from dtStudents

Each row is processed to generate a personalized salary slip

🔎 3. Check if Template Exists
Checks whether salary_template.docx exists in Downloads

If exists: Deletes an older version of TempSlip.docx in the invoice folder

If not exists: Does nothing

📝 4. Copy Word Template
Copies salary_template.docx to a new file TempSlip.docx

This is the working file where values will be inserted

✍️ 5. Fill the Word Template
Word Application Scope: Opens TempSlip.docx

Uses multiple Replace Text activities to substitute placeholders like:

<<EmployeeName>> → CurrentRow("EmployeeName")

<<Designation>>, <<BasicSalary>>, <<HRA>>, etc.

Replaces all placeholders in the document with the actual data from Excel

📄 6. Export as PDF
Exports the modified Word document as a PDF:

Filename: Salary_Slip<EmployeeName>.pdf

📧 7. Send Email
Sends the generated PDF to the employee’s email address:

Subject: "Your Salary Slip - <EmployeeName>"

Body: "Dear <EmployeeName>, Please find your salary slip attached."

Attachment: Salary_Slip<EmployeeName>.pdf

Uses SMTP with Integration Service for sending the email





<img width="601" height="844" alt="image" src="https://github.com/user-attachments/assets/aef9e936-478f-4c92-a734-a96026dafa7a" />

                             sample output:
