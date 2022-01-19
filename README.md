## How It Works

* Takes user input to determine what month should be analyzed and presented.
[User month and year input](imgs/month-input.png)

* Pulls transactions from user provided information (Example credit card transactions shown below).
[Credit card transactions](imgs/ccd-example.png)

* Goes through transactions to determine how they should be categorized
  * When reaching an unknown merchant, a category will be chosen and remembered. After choosing a category, the merchant name will autofill in the next input to see if the user wants to adjust the merchant name (For cases like chains using slightly different names for each store). 
  [Choosing a new merchant](imgs/new-merchant.png)
  
  * When reaching a merchant marked as manual, will prompt user to select a category (Useful for merchants that sell a wide variety of things).
  [Selecting a manual category](imgs/manual.png)
  
 * Prompts user to manually input balance information.
 [Entering balance information](imgs/balances.png)
 
 * Populates multiple worksheets in Excel with temporary or permanent information. Relevant permanent information shown below.
  * Data pulled for charts added to this worksheet. It will determine if an existing month is being analyzed and update that row if it already exists or it will add a new row.
  [Chart data example](imgs/CD-example.png)
  
  * Populate the Overview worksheet with information for the month and also save the same information to the Historical Data worksheet for records.
  [Overview example](imgs/overview.png)
  
  ## Dependencies
  
  * pyautogui
  * openpyxl
