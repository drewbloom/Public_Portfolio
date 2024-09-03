# Overview
This project was designed to automatically check institutional education standards for compliance with national standards.  To do this, a web application processes a file upload (pdf or Google Doc) containing the institutional standards.  It pushes the file text to be parsed by OpenAI's GPT-4o based on few-shot prompting methods to return a list of the learning objectives.  The learning objectives are added to a database (Google Sheets) and have embeddings created.  Those embeddings are compared to embeddings for the national standards, and a cosine similarity score of >0.50 is used to determine matches.  A series of functions within the sheet and within the Apps Script then create visuals and pinpoint relevant takeaways to be populated in a web application.  Finally, the Web Application returns these visuals and results of the report to the user's webpage once these processes finish on the back end.

### Situation
The CCO (Chief Curriculum Officer) for the organization wanted a workshop tool for use at an annual conference.  This conference acts as our main adoption and use event for each year.  This tool would allow medical education institutions to review their program objectives' alignment with national standards.  The goal was to have an easy-to-use tool for producing a report that could help our organization facilitate discussion and drive adoption of other features offered through our educational programming.

### Task
Produce a web application that would read a pdf upload, locate the educational institution's program objectives using an AI parser, compare them to national standards using an embeddings model, and produce data that the web application would use to create a report for each institution.

### Action
Met with the CCO to understand the situation and task, created a mock-up that would work for most inputs, reviewed the draft outputs with two institutional stakeholders and the CCO, revised the output and added new analysis for the data, and deployed a final web application ready for use in a conference workshop.

### Result
This project delivered an easier-to-use, more polished application than expected, leading to easy use and facilitated discussion of the report results.  It also demonstrated that simplified reporting tools for A&U could be easily built by existing team members, and our organization decided to allocate more of my labor resources to A&U feature enhancements.
