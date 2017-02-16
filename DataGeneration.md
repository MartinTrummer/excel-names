# excel-names: Data Generation

The VBA code that has been used to genearte the data on the 
[Excel Sheet](NameRulesUnicode64k.xlsm) is contained in the
[VBA module mExcelNamesGenerator](source/mExcelNamesGenerator.bas).

This was inspired by [this Stackoverflow anwser](http://stackoverflow.com/a/41877926/1041641) 
by user [Gary's Student](http://stackoverflow.com/users/2474656/garys-student).

The idea is that we simply iterate over all unicode characters and 
try to create an Excel Name for these 3 cases:
- *Character as name*: only the character is used to create the Excel Name: e.g. "a", "b", ...
- *Character at the start*: the character is used as the start of a name: e.g. "a_", "b_", ...
- *Switch*: backslash followed by a single character: e.g. "\a", "\0", ..
- *After the start*: the character is used after the start of an Excel Name: e.g. "_a", "_b", ..

When the `Names.Add` function returns an error (or a different name), we mark the Excel Name invalid - otherwise valid.  
Note: examples for a different name
  - `ActiveSheet.Names.Add "a ", " "` will create the name "a" (space is automatically removed)
  - `ActiveWorkbook.Names.Add ChrW$(173) & "_x", " "` might create the name "_x" (charcode 173 adn 1600 are automatically removed): see #2

# Performance
This chapter provides some info about performance improvements that have been made for the generation code.  
Still, the calculation is slow. 

Generating the worksheet names may take up to 10 minutes on an i7-2630QM CPU @ 2GHz with 8GB RAM.  
And during this time Excel (and even Windows) will not be responsive - and it may seem that it just hangs in an endless loop.  
So if you really want to start the generation routine on your PC, better plan a coffee break :)

When you also create the Workbook names, it takes about 3 hours (yes: hours)!

## Working with Excel Names is slow
Adding and deleting of many Excel Name in Excel 2013 is VERY slow.  
Thus we delete all the created Excel Names immediately. But also this is not enough.  
It seems that Excel does not really delete the Excel Names immediately. 
Thus, we create a temporary sheet and use it for 500 Excel Names. Then we delete the sheet and create a new one.

## Excel settings
Execution of the VBA code is much faster, when certain Excel features are disabled: e.g. ScreenUpdating, Automatic Formula Calculation, etc.


