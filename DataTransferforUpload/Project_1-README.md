Project 1 is a Google Apps Script macro designed to save hours of manual data entry.
It also helps the organization avoid human errors and pinpoint incompatibility errors between the upload file and web application.

Here is a link to the Google Doc for you to test - Make a copy for yourself to copy the script, then use the "Upload Sheet" custom menu located on the top toolbar to run the script:
https://docs.google.com/document/d/15JAJeGi8zTIsRZ5oqzHHdur8zsy_MTCRp460gh1htEU/edit?usp=sharing

# Situation:  
Team members in Assessment branch were spending hours each month copying and pasting fields for new questions authored in a shared Google Doc.
The team moved the text into one sheet to transpose it and edit the text to match formatting requirements for upload.
Then, the team uploaded the questions and often discovered additional unchecked errors.

# Task:  
Create a full automation of data transfer from collaboration document into an upload sheet for review and csv download

# Action:  
Using a total of 10 hours of work time, I met with Assessment team members to find out the needs of the project.
I was provided the three documents used in the current process and worked independently to create a demo of my solution.  I met again with one team member for initial feedback, then returned to a second team member for final approval and deployment. Total turn time was 2 work days with a total hour commitment by all parties equal to 10 labor hours.

# Result:  
Created a fully automated process to bypass transposing sheet and create upload document from Team Collaboration document, saving est. 10 hrs per month for Assessment Team staff

Below is the summary of functionality for the final product:

# Question Upload Automation
This script takes the data from a collaboratively authored document and formats it in a Google Sheet so that the data is instantly ready for final review, download as a csv, and upload to the organization’s Learning Management System (LMS).

It saves question authoring staff from copying and pasting almost every field of 70 tables (~90 pages) from this document into a Google Sheet manually (a time commitment of multiple hours and prone to errors)

The script locates table pairs by the strings “Question ID” and “Vignette”, then reads their contents, processes the text to meet upload requirements, and places each table pairing in a single row of the Upload Sheet used to enter this into the LMS.  It is left as a Google Sheet to allow Assessment staff to review any errors and make final checks prior to upload.

The script only inserts the question if its status is ‘Finalized’.  The script formats some text to reduce upload errors.  In that interest, it strips out bullet points, takes only Teaching Point titles, checks that the correct answer matches its partner string located in the distractor choices, and checks whether multiple learning objectives are found.  If the checks are not passed, the row is colored and the erroneous cell highlighted.