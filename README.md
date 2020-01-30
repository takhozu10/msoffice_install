# Microsoft Office update script for Mac

This script will assist with Microsoft Office update by identifying if the application is currently running on the computer.
If the application is not running, it will execute the command to run the update using JAMF Pro.

1. Create multiple Smart Computer Groups. Repeat the process for each MS Office apps (Excel, PowerPoint, Word, OneNote, Outlook)
  * Enter Display Name: MS Excel for Mac Update Group
  * For Criteria:
    * **Application Title** **has** **Microsoft Excel.app**
    * **and** **Application Version** **is not** **16.33.20011301** (or the latest version of MS Excel)
    * **and** **Application Version** **like** **16.** (This portion prevents from running against MS Excel 2011)
2. Create new policy
  * Enter Display Name: Microsoft Office update script
  * Trigger: Recurring Check-in
  * Execution Frequency: Ongoing
  * Add Scrip: [msoffice_update.rb](msoffice_update.rb) script
  * Scope Targets add all smart groups created in step 1.
3. Create new policies for installing each app. Repeat the process for each MS Office app (Excel, PowerPoint, Word, OneNote, Outlook)
  * Enter Display Name: MS Excel for Mac update
  * Trigger: Custom
  * Custom Event: ms_excel_update
  * Execution Frequency: Ongoing
  * Add Microsoft Excel Update Package
  * Scope: MS Excel for Mac Update Group (Add Limitation to prevent update via DMZ)
4. At this point the policy created in Step 2 will run the script every 15 minutes. Each time, if respective app is not running, it will trigger the respective app update package to run. 
