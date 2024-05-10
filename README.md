# Health Services Scheduler
## This program should import an excel spreadsheet format it and import it to an SQL Database
---
Currently working on the logic to exclude the breaks from the import. I'm having a bit of an issue selecting the specific shifts, currently it is looking for shifts which time === 30 minutes. 
## Update:
Used a Regex to extract the specific information from the spreadsheet, added a function to calculate total shift in minutes, filtered out the shifts that are === to 30 minutes, then added 30 minutes to any the remaining shifts (still need to exclude shifts that are < than 6 hours since those wont have a break on them, but that shouldnt be too hard). The order in which the function is applied is important.
