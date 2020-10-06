# Outlook-Email-Generator
### Pre-Introduction Note
This code was created for the automation of incident and task related emails. However, this can easily be changed by removing the section of the form that is with regards to incident and task numbers. Though refactoring the position of the buttons may be wanted then. 

## Introduction
This is a PowerShell script that allows for the generation of emails through HTML templates stored in the folder Email__Templates to be put directly in the drafts folder of Outlook. Creates a UI for the user to type in their ticket number to add to the email if need be. 

The problem this program looks to solve is making it easier for users to send template emails, and allow for unity in the emails that are sent through groups. It allows for new templates to be created and distributed easily in case of updates or loss of requirements. 

## How to Use
1. Make sure that powershell can be run on your computer
2. Create an email template using the email template html file
3. Run the script through powershell

## Notes
 - The code automatically ignores the email_template file and therefore that file can remain in the email_templates folder

 - The code automatically connects to outlook through your existing outlook connection and doesn't need additional signing in

 - A shortcut can be created for the script to run more smoothly, would recommend adding this line to the target field:  <br>  -ExecutionPolicy Bypass -WindowStyle Hidden -File

 - The code automatically will expand the UI with regards to the number of templates in the email_templates folder. 

