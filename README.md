# ticket_prepper
A script to prepare the tickets for edgeconnex locations. Currently only supports AMS01.

To run the ticket prepper:
- Export the tickets from edgeos.edgeconnex.com, while you're logged in.
- Make sure the file extension is .xlsx, other extensions are not supported. 
- Run ticket_prepper.exe
- The program does the following:
	* Lets you select the file that contains the tickets that need to be opened.
	* Opens each and every ticket that is relevant for the SOC in a seperate chrome window.
	* Returns a file with relevant tickets, although one could use the original file.
      
**Possible Future Updates**

Features that might be added:
- Automatically updating the tickets with the correct responses.
- Automatically color the excel file.
- Creating an extra .xlsx file that will generate a list of future tickets.
- Generate a list of tickets that should be checked extra.
