## Overview of recharge model

The recharge model has been developed, which can be summarized as follows:
* To ensure smooth facility operations, monthly payments determined by user type will be charged. I've assigned lab groups to user types [see here](https://docs.google.com/spreadsheets/d/1Yom-H6j04TJ5_W1Ic2dlqKsVWrHQbdInQyq_Q7ZVnZA/edit?usp=sharing) based off prior information given to me, but groups may choose to change types. The"User types" (Core, Associate, Regular) described below:
  * Core - $650 / month; equipment use costs will be multiplied by 1
  * Associate - $175 / month; equipment use costs will be multiplied by 1.5
  * Regular - $0 / month; equipment use costs will be multiplied by 2
* Large purchases (e.g. bulk order of lab consumables, equipment servicing, etc.) are effectively amortized over the fiscal year. Groups are charged for only the consumables they use (e.g. purchased screens, Mosquito / Dragonfly supplies). Equipment service fees will be, whenever possible, paid in monthly installments. 
* Active lab groups (those that will be charged) for that month are determined by the usage logs of the RockImagers, Mosquito dropsetters, and Dragonfly.
* In the last month of the fiscal year, a final adjustment credit or recharge will be made to carry a small surplus to the next year or zero the yearly balance, respectively. Based off projections, lab groups will most likely be credited the proportional amount overpaid.  

The main takeaways of these changes are:
* A more equitable, evenly distributed, and transparent recharge model; lab groups can find itemized spreadsheets detailing their facility usage and expenses [here](https://drive.google.com/drive/folders/1DvilgbyLRI9At3wwR9F2D4Dufn77VOU6?usp=sharing)
* An attempt to minimize financial risk and facility downtime
* Most of the recharge process and calculations are automated (reduce human error)

---

## RechargeSummary spreadsheet explained

* **Facility Fee //** Facility Fee of your lab's  User Type 

* **Screens Total Costs //** Total cost of ordered crystallization screens. 

* **Mosquito Total Cost //** The total cost of Mosquito usage. See `[your_lab]` -> `mosquitoUsage[date_range].xlsx`

* **RockImager Total Cost //** The total cost of RockImagers usage. See `[your_lab]` -> `20c_rockImagerUsage[date_range].xlsx` and `[your_lab]` -> `4c_rockImagerUsage[date_range].xlsx`

* **DFly Total Cost //** The total cost of Dragonfly usage. See `[your_lab]` -> `dragonflyUsage[date_range].xlsx`

* **Raw Usage //** The raw usage (min) of all equipment taken from log files. 

* **Usage prop //** The sum of your lab's usage divided by all lab's usage.

* **Month Dist. Cost //** Non amortized facility expenses (liquid N2, computing licenses, etc) calculated by `Usage prop` x `Use Multiplier` x `Total distributed cost`

* **Payroll Cost //** If `Total payroll cost` > `Total facility fee`, `(Total payroll cost - total facility fee)` x `Usage prop.`Else, cost is 0.

* **Use Multiplier //** Use Multiplier of your lab's User Type.

* **Total Charge//** `(Month Dist. Cost + Mosquito Total Cost + RockImager Total Cost + DFly Total Cost)` x `Use Multiplier` + `Screens Total Cost` + `Payroll Cost`

---

## How-to procedure (for internal use):
Recharge is to be done monthly. Finance department (La-Risa Lewis, La-Risa.Lewis@ucsf.edu) will send an Excel spreadsheet detailing the facility's payroll (DPE) and purchases (GL) for the prior month. 
* [1] Get DPE and GL files
  * Upload Excel spreadsheets to `My Drive > ljiang > xrayFacilityRecharge > GL_DPE > excelOriginals` and rename with format `[month]_[year] [GL/DPE]`
    * File name and location are important various scripts. 
    * In the `GL_DPE` folder, open the Google sheet `collatedGL_DPE` and under `Custom functions` run `Make copies from Excel`. This will create Google sheet copies in the  `sheetsCopies` folder.
    * Run the custom function `Collate itemized recharge`. If there is an error at this step, it will mostly be due to how certain column headers are named. You might need to rename certain columns to match the format of existing Google sheets.
    * Note, although not recommended, you can also directly enter the GL and DPE data into `collatedGL_DPE` Google sheet, but be sure to format cell values accordingly.
* [2] Get 4C and 20C RockImager logs
  * Go to each imager's computer and get the event log file by:
    * `File` --> `Reports` --> `Event log`
    * Specify the date range for the month of interest (e.g. Aug 2018 would be 8/1/2018 and 9/1/2018).
   * Appropriately name the .txt which should ideally comma-delimited and put into corresponding folder `My Drive > ljiang > xrayFacilityRecharge > equipmentLogs > RockImagerEventLogs > rockImager 4C / rockImager 20C`
* [3] Run python script
  * Run the python script `calculateRecharge.py` (see section "calculateRecharge.py details" for notes) from cmd or terminal and enter date range desised. For e.g., if the recharge is for July 2018, you would enter 2018-7-1 as the start date and 2018-8-1 as the end date.
  * If successful, the script will take usage logs from the Mosquito(s), RockImagers, Dragonfly, and online screen orders and spit out itemized spreadsheets per lab in a directory named accordingly to the date range entered.
  * Take this folder and upload it to the parent folder named `xrayFacilityRecharge > monthlyRecharge`.
  * Lastly, send the `rechargeSummary[date_range].xlsx` to the finance dept. (La-Risa.Lewis.ucsf.edu). 

## calculateRecharge.py details:
* Python (ver. 3.6.3) packages used (as of 6/12/19) -- may work on other versions as well. If missing anything, then a pip install should do the trick. 
  * pandas (ver. 0.23.4)
  * pygsheets (ver. 2.0.1)
  * numpy (ver. 1.5.4)
  * oauth2client (ver. 4.1.3)

* The script needs to take in `msg-Recharge-24378e029f2d.json` in order to authorize access to the Google sheets needed. Make sure this file exists; otherwise, you will have to generate a new json file.
