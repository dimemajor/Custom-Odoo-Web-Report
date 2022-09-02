# Custom-Odoo-Report
This script was created to customize sales report from odoo point of sales. 

The default report that comes out from odoo point of sale report is a summarized version of what was sold.
For example, lets assume you have in stock a bag with 2 variants (colours in this case), say, black and blue. Lets also assume you have a bottle with 2 variants (white and red). If in a day you sold:

2 blue bags at 5000
3 black bags at 7500
1 white bottle at 600
5 red bottles at 3000

The odoo report comes out as:
bags 12500
bottle 3600
Total 13100

This summariced format is not very useful if you want to be able to see the exact variant that was sold. It would be nice to have a report that does that right? Luckily, odoo allows customized modules where this can be achieved but that comes at a price. A relatively huge price increase for that feature, especially if that is the only customization you need. Which is where this script comes in.


In this instance, this script sets to:
-grab variants sales data and method in which the payments were made from odoo website and essentially,
-stitches the data to their respective local pictures of the products on the computer it is run from
-writes that data to a word document and attaching corresponding links to the local pictures on the computer and,
-converting that document to a pdf report


In the program, 
-An interface is provided for the user to select the date and time interval for which the report to be generated is. A default of last start pos session to current date and time is given. 
-When the 'print' command is given. The program logs into the website and saves the cookies/session_id for future use.
-Makes proper post requests (odoo uses apis) to get all the necessary data, does necessay cleanups and loads the data into appropriate dictionaries and lists
-Looks for the folder containing the pictures of the products sold (the name of the products in odoo and pictures must match)
-Writes into word document in a tabular format while providing links to the files in the system
-Converts the docx to pdf


Happy coding












