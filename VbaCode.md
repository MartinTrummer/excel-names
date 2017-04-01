# excel-names: VBA code

All VBA code in this project is in the public domain (see <http://unlicense.org/>). 
So feel free to copy the code to your own VBA projects and change it to your needs.

If you find any issues or have ideas to improve it, please 
file [an issue](https://github.com/MartinTrummer/excel-names/issues).

All the code is included in the [Excel Sheet](NameRulesUnicode64k.xlsm).  
For convenience all VBA modules of the [Excel Sheet](NameRulesUnicode64k.xlsm)
have been exported to separate files in the [source directory](source/), 
so that you can directly check the code without Excel and also because we 
can use the plain text files with the version control system.

## Prepare your VBA project
Simply copy the code of the following files to your VBA project:
- [mExcelNameRulesData.bas](source/mExcelNameRulesData.bas):  
  This file just contains data about unicode characters and no meaningful VBA functions.
- [mExcelNameRules.bas](source/mExcelNameRules.bas):  
  This file depends on the [mExcelNameRulesData.bas](source/mExcelNameRulesData.bas) 
  and it contains functions that you can use in your VBA code.

## Excel Name Rules
The Microsoft documentation is not very clear about naming rules 
(see [Learn about syntax rules for names](https://goo.gl/k4Ne1E)). 
This document describes in details rules that are used in the VBA code.

### Character Check Functions
For Excel Names that consist of many characters we cannot test all possible permutations.  
So the results may return wrong results in edge-cases: 
if so, please file [an issue](https://github.com/MartinTrummer/excel-names/issues).

- `Names_IsCharValidAsName(sCharacter As String) As Boolean`  
  Since the generator script tried all of these characters, you can be sure of the result:
  When this function returns `true`, you can use the character as an Excel Name, otherwise not.  
  Examples:
  - `Names_IsCharValidAsName("a")` returns `true`
  - `Names_IsCharValidAsName("c")` returns `false` (c is invalid)
- `Names_IsCharValidAtStart(sCharacter As String) As Boolean`  
  When this function returns `true`, you can probably use the char at the start of an Excel Name.  
  When it returns `false`, you can for sure not use it.  
  Examples:
  - `Names_IsCharCodeValidAtStart("a")` returns `true`
    - "ax" is okay
    - "\a" is not okay (a switch)
    - "a$" is not okay ($ is always invalid)
    - "a1" is not okay (it is a cell-reference)
  - `Names_IsCharCodeValidAtStart("?")` returns `false`
    - "?a" is not okay (? is invalid at the start)
- `Names_IsCharValidAfterStart(sCharacter As String) As Boolean`     
  When this function returns `true`, you can probably use the char after the start of an Excel Name.  
  When it returns `false`, you can for sure not use it.  
  Examples:
  - `Names_IsCharValidAfterStart("?")` returns `true`
    - "a?" is okay (note: "?", "?a" are not okay)
  - `Names_IsCharValidAfterStart("$")` returns `false`
    - "a$" is not okay ($ is never okay )

### Excel Name Check Functions
- `Names_IsValidName(sNameToTest As String) As Boolean`  
   Check if the name is valid:  
  - `true`: Excel name is probably valid
  - `false`: Excel name is for sure not valid
- `Names_AdjustName(sNameToTest As String, Optional sReplaceChar As String = "_") As String`
  - returns a string that is probably a valid an Excel Name

Rules for adjusting an Excel Name ("_" as replace character):
- Blank: "" -> "_"
- Invalid single character: convert to replace char
  - "$" -> "_"
  - "!" -> "_"
- Invalid single start character - prepend replace char
  - "c" -> "_c"
  - "1" -> "_1"
- Invalid start char but valid afterwards - prepend replace char:
  - "?x" -> "_?x"
- Switches (backslash followed by a single character) - prepend replace char:
  - "\a" -> "_\a"
- Invalid start char (and also valid afterwards) - converted to the replace char:
  - "$x" -> "_x"
- Spaces are converted to the replace char:  
  Note: when you use VBA to set an Excel Name that has spaces at the start or end, the spaces will be trimmed automatically.
  - " " -> "_"
  - " a" -> "_a"
  - "a " -> "a_"
  - " a " -> "\_a\_"
- Too long identifier: the characters at the right will be cut  
  Excel 2013 allows a max length of 255 characters
- Cell-ref like: prepend the replace char
  - "A1" -> "_A1"
  - "R1C1" -> "_R1C1"
  - Note: this check is intentionally more strict than necessary: 
    - Switches (e.g. "\a") are generally disallowed, because they are all invalid on the Workbook.
    - see ["Cell-References"](#cell-references) below for details


### Cell-References
The Excel Name check functions in our VBA code are overly strict for cell-reference-like Excel Names.  
This is intended to keep the check-code simpler and to improve upwards-compatibility to future Excel versions.  
We **disallow ALL** Excel Names that start with letter, followed by one or more digits (which applies to all the examples in the table below).

Notes:
- "R1048577C1" is a valid Excel Name (in Excel 2013)  
    - We guess, because the row 1048577 is higher than the max. possible row in Excel 2013.  
    Future Excel version may allow more rows and then this reference may become invalid in those versions.  
    Thus we decide to treat this Excel Name as invalid.
- There are also some variations that are invalid Cell-References, but still cannot be used as a Name.  
  e.g. "R1C16385xxx" is not a valid cell-reference (note the trailing "xxx" characters after the R1C1 notation), 
  but still you cannot use this as an Excel Name.

Excel 2013 Examples (X = INVALID):

| Example | Cell-Ref | Excel-Name | Note |
| --- | --- | --- | --- |
| R1048576C1 | VALID | X | R1C1 reference (row lower than max-Excel rows) |
| R104857**7**C1 | X | VALID | R1C1 reference (row higher than max-Excel rows) |
| R1C16384 | VALID | X | R1C1 reference (col lower than max-Excel cols) |
| R1C1638**5** | X | X | R1C1 reference (col higher than max-Excel cols) |
| R1C | X | X | not an R1C1 reference (col is missing), still invalid as an Excel Name |
| R1D | X | X | not an R1C1 reference, still invalid as an Excel Name |
| R1C16385**xxx** | X | X | invalid R1C1 reference because there are letters after the col |
| A1 | VALID | X  | valid A1 notation  |
| XFD1048577 | VALID | X  | valid A1 notation (max row/col) |
| XFD104857**8** | X | VALID | invalid A1 notation, but can be used as an Excel Name |
| XF**E**1048577 | X | VALID | invalid A1 notation, but can be used as an Excel Name |

### Miscellaneous Information
- Character-Case   
  Keep in mind, that Excel Names are case-insensitive.  
  So if you create an Excel Name like "abc" and then another Excel Name "ABC", the later Excel Name will overwrite the first one.
- Edge-cases  
  There are for sure some edge cases that are currently not covered by the VBA code.  
  If you find any issues or have ideas to improve it, please 
  file [an issue](https://github.com/MartinTrummer/excel-names/issues).  
