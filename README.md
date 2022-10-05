# Custom-Odoo-Report
## Summary
This program was created to customize sales report from odoo point of sales web app. 

The default report that comes out from odoo point of sale report is a summarized version of what was sold.
For example, lets assume you have in stock a bag with 2 variants (colours in this case), say, black and blue. Lets also assume you have a bottle with 2 variants (white and red) and there are 5 bottles in a pack. If in a day you sold:

- 2 blue bags at 5000 each
- 3 black bags at 7500 each
- 1 white bottle at 600 each
- 11 red bottles at 700 each

The odoo report comes out as:
| Product | Quantity | Price Unit |
| --- | --- | --- |
| Bag | 5 pce | 6250 |
| Bottle | 12 pce | 650 |

The aim of the program is to have the report in this format instead;

| Product | Quantity | Price Unit | Total Price | Image
| --- | --- | --- | --- | --- |
| Bag (Black) | 2 pce | 5000 | 10000 | ![gdrive link](https://google.com)
| Bag (Blue) | 3 pce | 7500 | 22500 | ![gdrive link](https://google.com)
| Bottle (White) | 1 | 600 | 600 | ![gdrive link](https://google.com)
| Bottle (Red) | 2 pck 1 pce | 700 | 7700 | ![gdrive link](https://google.com)

** The image column is also accompanies a thumbnail of the photo it is linked to. **

## How it works:
The program sets to:
-grab products, their variants, sales data, and method in which the payments were made from odoo different areas of odoo website/database/api and essentially,
-stitch corresponding data to their respective local pictures (of the same name with the product themselves) of the products on the computer it is run from
-writes that data to a word document and attach corresponding links from google drive and,
-converts that document to a pdf report


In the program, 
-An interface is provided for the user to select the date and time interval for which the report to be generated is. A default of last start pos session to current date and time is given. 
-When the 'print' command is given. The program logs into the website and saves the cookies/session_id for future use.
-Makes proper post requests (odoo uses apis) to get all the necessary data, does necessay cleanups and loads the data into appropriate dictionaries and lists
-Looks for the folder containing the pictures of the products sold (the name of the products in odoo and pictures must match)
-Writes into word document in a tabular format while providing links to the files in the system
-Converts the docx to pdf

## Extensibility
This program is created specifically for a client but is fairly reusable for another potential user. It will however need some changes especially to the "constant.py" file (not included in the repo for its sensitivity) which include the following variables;
- DOMAIN (eg. https://test.odoo.com)
- EMAIL (user email)
- PASSWORD (user password)
- FOLDER_NAME (name of folder where report files are saved)
- COMPANY (name of company for customization of pdf)
- CATALOG_FOLDER (name of folder where pictures are located)
- SCOPES (g_drive scopes if integrating with google drive)













