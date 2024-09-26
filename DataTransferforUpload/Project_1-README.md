Project 1 is a Google Apps Script macro designed to save hours of copyediting and manual data entry.
It also helps the organization avoid human errors and pinpoint incompatibility errors between the upload file and web-hosted learning management system.

Here is a link to the Google Doc for you to test - Make a copy for yourself to copy the script, then use the "Upload Sheet" custom menu located on the top toolbar to run the script:
https://docs.google.com/document/d/15JAJeGi8zTIsRZ5oqzHHdur8zsy_MTCRp460gh1htEU/edit?usp=sharing

# Situation:  
Team members in Assessment branch were spending hours each month copying and pasting fields for new questions authored in a shared Google Doc.
The team moved the text into one sheet to transpose it and edit the text to match formatting requirements for upload.
Then, the team uploaded the questions and often discovered additional unchecked errors.

# Task:  
Create a full automation of data transfer from collaboration document into an upload sheet for review and csv download.  In addition, add programmatic copyediting to eliminate error triggers or style errors from making their way into the upload document

# Action:  
As the lead on this project, I met with Assessment team members to find out their needs from the automation and sampled current SOPs.
I was provided the three documents used in the current process and worked independently to create a demo of my solution.  I met again with one team member for initial feedback, then returned to a second team member for final approval and deployment. After initial deployment, team members requested an additional copyediting layer, which was completed with a return to the same process of information gathering, test creation, feedback, revision, and deployment.

# Result:  
Created a fully automated process to bypass transposing sheet and create upload document from Team Collaboration document, saving est. 10 hrs per month for Assessment Team staff

Below is the summary of functionality for the final product:

# Copyediting and Question Upload Automation
This script takes the data from a collaboratively authored document and corrects style issues and error triggers in a specific table edited by multiple authors and prone to errors.  After copyediting is completed, the document is reviewed by a team for final approval.  The upload automation then formats it in a Google Sheet so that the data is instantly ready for final review, download as a csv, and upload to the organization’s Learning Management System (LMS).

It saves question authoring staff from performing manual copyedits before copying and pasting almost every field of 70 tables (~90 pages) from these documents into a Google Sheet manually (a time commitment of multiple hours and prone to errors).  Deployment in an Apps Script gave Assessment team members the transferability needed to locate these scripts in all question writing documents, which are created regularly as needed and exist at varying levels of review and completion at any given time.

The copyedit script only makes edits on human-edited tables, which contain "Vignette:" in their first cell.  It proceeds through a series of functions for the entire table and for specific cells located within the table to eliminate common errors found in these documents.  The upload script locates table pairs by the strings “Question ID” and “Vignette”, then reads their contents, processes the text to meet upload requirements, and places each table pairing in a single row of the Upload Sheet used to enter this into the LMS.  It is left as a Google Sheet to allow Assessment staff to review any errors and make final checks prior to upload. The script only inserts the question if its status is ‘Finalized’.  The script formats some text to reduce upload errors. It also makes a series of checks to determine whether the row contents match the expected contents when passed from Docs to Sheets.  If the checks are not passed, the row is colored and the erroneous cell(s) highlighted.
