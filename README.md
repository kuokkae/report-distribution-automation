# report-distribution-automation
if you pay someone to sit in front of their laptop everyday just to refresh your excel report that is connected to your database, you might want to read this

# purpose:
to automate the mundane tasks of back-up, refresh, save file and copy daily

# approach:
Excel report are pre-built. The work is split into these phases
1. Backing-up (powershell)
2. Refreshing (powershell)
3. Distribution (powershell)
4. Scheduling (batch file)

Macro loaded reports are prone to error, and backing up the previous day file can save time in trying to rebuild the file when your copy is corrupted. The report then launches and refresh the whole file, including any macro that needs to run, saves and closes. Distribution consist of copying from local drive to a shared folder, in our case we used OneDrive. Finally scheduling is done using batch file that launches the powershell script in succession.

# benefit
this shaves 1 hour a day simply sitting in front of the laptop to perform these tasks, and they are done with high precision with zero wait time between each task unlike human interactions
