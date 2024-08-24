Project 3 is a data consolidation project completed for use by internal users with varying levels of permissions.
You can see the simple web UI by following this link:
https://script.google.com/macros/s/AKfycbxlK4TRcVwkyYbs2tqTxfRAla8hOzVyNKa6JbzP4lo/dev

And you can request to see a scrubbed copy of the database if you want to see any of its actual functionality.

# Situation:
The organization's current metadata downloads did not give the educational organization's curriculum team enough data to easily review and author new content without engaging in time-consuming manual searches across numerous webpages.  Some metadata existed that was useful, while other data was inaccessible via downloads.

Our dev team had a tight deadline for other feature adds, and they were weren't able to process minor bug fixes or data requests for the back half of the quarter - a request for this data through our own data warehousing was left unanswered for over a month.  Because the team needed this data extracted and consolidated, it was decided that webscraping was the best approach.

# Task:
Combine existing, easily-accessible data with additional webscraped data into one database, allowing curriculum team members to query data as needed for reviewing and authoring projects.
In addition, create a web interface that would mask all of the utilities of the database and allow internal users with fewer permissions to access some data.

# Action:
Created a scrape-first, parse-later approach to pull the entire library of content from our LMS, read existing database content, and parse to locate and append the additional data to its corresponding rows in the database.
Created a web interface to render some basic query functions in a webpage, masking access to the main database content for more limited users.

# Result:
The database queries saved hours of manual data collection for new content authoring and content review processes.  Team members were extremely pleased with the ease of use.
The webpage UI allowed the project to be expanded beyond the direct 7-person team and used as a limited data search tool by the entire organization and organizational affiliates involved in the authoring and review of educational content.
