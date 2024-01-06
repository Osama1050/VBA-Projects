# Undo after running a VBA code
this is the functional code to make the undo button function for a number of times after running a VBA code.
this solves the problem that arrise whenever running a VBA code that make changes to the workbook, which is clearing undo cache.

# Instructions
- open the .xlsm file.
- you will find the code in Sheet1 module.
  (this will work in the sheet that has the code in it only!)
- you can adjust the public variable "ChangeLimit" to control how many undo times.
- close VBA editor, and start entering data to sheet.
