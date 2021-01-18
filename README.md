# ticket_prepper
A script to prepare the tickets for edgeconnex locations.

To run the ticket prepper:
- Export the tickets from edgeos.edgeconnex.com, while you're logged in.
- Change the name of the .xlsx file to: "input_tickets.xlsx".
- Run ticket_prepper.exe
- The script will do two things: 
      * Open a browser window which will have all the to be checked tickets opened.
      * Return the excel file: "tickets_DATE.xlsx" which contains solely the tickets concerned to SOC.
      
**Future updates**

Features that might be added:
- Automatically updating the tickets with the correct responses.
- Automatically color the excel file.
- Creating an extra .xlsx file that will generate a list of future tickets.
- Generate a list of tickets that should be checked extra.
