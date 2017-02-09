Attribute VB_Name = "mExcelNamesGenerator"
'
' This is free and unencumbered software released into the public domain.
'
' For more information, please refer to <http://unlicense.org/>
' or to the UNLICENSE text file contained in this repository
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
' EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
' MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
' IN NO EVENT SHALL THE AUTHORS BE LIABLE FOR ANY CLAIM, DAMAGES OR
' OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE,
' ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR
' OTHER DEALINGS IN THE SOFTWARE.
'
Option Explicit

' NOTES:
' chr 32  spaces at the start and end will be removed automatically
' case is ignored: zABC = Zabc
' chr 173 at the start cannot be entered in the name box (left top) of the spreadsheet, but in the name manager
'         char alone or after start works...


' see also:
' * Microsoft Online Help: https://goo.gl/PsGUQj
' * Stackoverflow: https://goo.gl/vHOIbL

' just select an identifier which is always valid (also at the start)
Private Const RANDOM_VALID_IDENTIFIERS = "X"

' see Chr help: https://msdn.microsoft.com/en-us/library/office/gg264465.aspx
Private Const NAMES_MAX_UNICODE_CHARACTER_CODE = 65535

Private Const NO_OF_HEADER_ROWS = 2

Private Const COL_CHAR_CODE = 1
Private Const COL_CHAR = 2
Private Const COL_CHAR_OTHER_CASE = 3
Private Const COL_CHAR_AS_NAME_IS_OKAY = 4
Private Const COL_CHAR_AT_START_IS_OKAY = 5
Private Const COL_CHAR_AFTER_START_IS_OKAY = 6

Private startDateTime As Double

Private Sub Dbg(sMessage As String)
    Debug.Print Now & ": " & sMessage
End Sub

' returns a range for the given column which starts
' at the row after the header rows and has
' 64k (NAMES_MAX_UNICODE_CHARACTER_CODE) rows
Private Function GetColRangeWithoutHeader(iCol As Long) As Range
    Dim oRangeStartCell As Range
    Dim oRangeEndCell As Range
    
    Set oRangeStartCell = Cells(NO_OF_HEADER_ROWS + 1, iCol)
    Set oRangeEndCell = Cells(NO_OF_HEADER_ROWS + NAMES_MAX_UNICODE_CHARACTER_CODE, iCol)
    
    Set GetColRangeWithoutHeader = Range(oRangeStartCell, oRangeEndCell)
End Function

' sets the first format condition to have a
' dark red font and a light red background
Private Sub FormatWarning()
    With Selection.FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
End Sub

' sets the first format condition to have a
' dark green font and a light green background
Private Sub FormatOkay()
    With Selection.FormatConditions(1).Font
        .Color = -16752384
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13561798
        .TintAndShade = 0
    End With
End Sub

' will format the given column green if the cell values are true
' red otherwise
Private Sub ConditionalBooleanFormat4Col(iCol As Long)

    Const IS_TRUE_STRING = "=TRUE"
    Dim oRange As Range
    Set oRange = GetColRangeWithoutHeader(iCol)
    oRange.Select
    
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlNotEqual, _
        Formula1:=IS_TRUE_STRING
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    FormatWarning
    Selection.FormatConditions(1).StopIfTrue = False
    
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:=IS_TRUE_STRING
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    FormatOkay
    Selection.FormatConditions(1).StopIfTrue = False
End Sub

Private Sub CondFormatWarnIfNotEmpty(lCol As Long)
    ' lower-case: highlight red if not the same as original
    Dim oRange As Range
    Set oRange = GetColRangeWithoutHeader(lCol)
    oRange.Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlNotEqual, _
        Formula1:="="""""
    FormatWarning
    Selection.FormatConditions(1).StopIfTrue = False
End Sub

' deletes all conditional formatting from the active sheet
Private Sub ClearAllConditionalFormats()
    ActiveSheet.Cells.FormatConditions.Delete
End Sub

' will apply conditional formatting to the generated columns
Private Sub CondFormat()
        
    ClearAllConditionalFormats
    
    CondFormatWarnIfNotEmpty COL_CHAR_OTHER_CASE
    ConditionalBooleanFormat4Col COL_CHAR_AS_NAME_IS_OKAY
    ConditionalBooleanFormat4Col COL_CHAR_AT_START_IS_OKAY
    ConditionalBooleanFormat4Col COL_CHAR_AFTER_START_IS_OKAY
    
    ' "unselect"
    Range("A1").Select

End Sub

' returns if it is possible to create the given name in the given worksheet
Private Function IsNameValid(sNameToTest As String, _
 Optional oWsTemp As Worksheet = Nothing) As Boolean
    Dim oNameObj As Name
    Dim bResult As Boolean
    Dim sSheetRef As String
    Dim sExpectedName As String

    bResult = False
    On Error Resume Next
    
    Err.Clear
    
    If oWsTemp Is Nothing Then
        Set oNameObj = ActiveWorkbook.Names.Add(sNameToTest, " ")
    Else
        Set oNameObj = oWsTemp.Names.Add(sNameToTest, " ")
        sSheetRef = oWsTemp.Name & "!"
    End If
    If Err.Number = 0 Then
        ' e.g. when you enter a name that ends with a space
        '      character "abc " the trailing space will automatically
        '      be ignored: i.e. the generated name is "abc" (not "abc ")
        '      and thus, the space character at the end is invalid
        sExpectedName = sSheetRef & sNameToTest
        If oNameObj.Name = sExpectedName Then
            bResult = True
        Else
            Dbg "mismatch: >" & oNameObj.Name & "< vs. >" & sExpectedName & "<"
        End If
       ' delete immediatley, because Excel cannot handle hundreds of names very well
       oNameObj.Delete
    End If
    
    On Error GoTo 0
     
    IsNameValid = bResult
    
End Function

' returns the other case for the given char or blank if
' the character has no upper/lowercase pendant
' Examples
'  GetOtherCase("a") --> "A"
'  GetOtherCase("A") --> "a"
'  GetOtherCase("1") --> ""
Private Function GetOtherCase(sCurrentChar As String) As String
    Dim sLcase As String
    Dim sUcase As String
    
    sLcase = LCase$(sCurrentChar)
    sUcase = UCase$(sCurrentChar)
    
    If (sLcase = sUcase) Then
        GetOtherCase = ""
    ElseIf (sLcase = sCurrentChar) Then
        GetOtherCase = sUcase
    Else
        GetOtherCase = sLcase
    End If
End Function

Private Sub GenerateData(Optional lStart As Long = 1 _
    , Optional lEnd As Long = NAMES_MAX_UNICODE_CHARACTER_CODE)

    Dim i As Long
    Dim sNameToTest As String
    Dim iCol As Long
    Dim oNameObj As Name
    
    Cells(NO_OF_HEADER_ROWS, COL_CHAR_CODE).Value = "Chr-Code"
    Cells(NO_OF_HEADER_ROWS, COL_CHAR).Value = "Char"
    Cells(NO_OF_HEADER_ROWS, COL_CHAR_OTHER_CASE).Value = "Other Case"
    Cells(NO_OF_HEADER_ROWS, COL_CHAR_AS_NAME_IS_OKAY).Value = "OK"
    Cells(NO_OF_HEADER_ROWS, COL_CHAR_AT_START_IS_OKAY).Value = "OK at start"
    Cells(NO_OF_HEADER_ROWS, COL_CHAR_AFTER_START_IS_OKAY).Value = "OK after start"
    
    Dbg "preparing calculation"
    
    Dim arrCalculated(1 To NAMES_MAX_UNICODE_CHARACTER_CODE, COL_CHAR_CODE To COL_CHAR_AFTER_START_IS_OKAY) As Variant
    Dim arrDuplicates(1 To NAMES_MAX_UNICODE_CHARACTER_CODE) As String
    
    Dbg "starting calculation"
    
    Dim oWsTemp As Worksheet
    Set oWsTemp = ActiveWorkbook.Worksheets.Add
    For i = lStart To lEnd
    
        If i Mod 100 = 0 Then
            Dbg "i=" & i
        End If
        If i Mod 500 = 0 Then
            ' note: if we used only one worksheet then creating/deleting
            '       names takes forever...
            Dbg "NEW worksheet " & i
            oWsTemp.Delete
            Set oWsTemp = ActiveWorkbook.Worksheets.Add
        End If
        
        Dim sCurrentChar As String
        sCurrentChar = ChrW$(i)
        
        arrCalculated(i, COL_CHAR_CODE) = i
        arrCalculated(i, COL_CHAR) = sCurrentChar
        arrCalculated(i, COL_CHAR_OTHER_CASE) = GetOtherCase(sCurrentChar)
        
        ' check character alone
        sNameToTest = sCurrentChar
        arrCalculated(i, COL_CHAR_AS_NAME_IS_OKAY) = IsNameValid(sNameToTest, oWsTemp)
        
        ' check character at the start
        sNameToTest = sCurrentChar & "_" & RANDOM_VALID_IDENTIFIERS
        arrCalculated(i, COL_CHAR_AT_START_IS_OKAY) = IsNameValid(sNameToTest, oWsTemp)
        
        ' check character AFTER the start
        sNameToTest = RANDOM_VALID_IDENTIFIERS & "_" & sCurrentChar
        arrCalculated(i, COL_CHAR_AFTER_START_IS_OKAY) = IsNameValid(sNameToTest, oWsTemp)
    Next i
    oWsTemp.Delete
    Set oWsTemp = Nothing
    
    Dbg "writing data to sheet.."
    
    Dim iStartRow As Long
    iStartRow = 1 + NO_OF_HEADER_ROWS
    Dim sStartCell As String
    sStartCell = "A" & iStartRow
    Dim oStartCell As Range
    Set oStartCell = Range(sStartCell)
    oStartCell.Resize(UBound(arrCalculated, 1), UBound(arrCalculated, 2)).Value = arrCalculated
    
    Dbg "creating named col-defines..."
    
    ' set dataXXX names
    Dim lNoOfRows As Long
    lNoOfRows = lEnd - lStart + 1
    CreateDataNames oStartCell.Resize(lNoOfRows), lNoOfRows
End Sub

' will create named ranges for all generated data
' the range will cover only the data rows for each column
' e.g. "dataCharCode" range is A3:A65537
Private Sub CreateDataNames(oFirstDataRange As Range, lNoOfRows As Long)
    SetDataName oFirstDataRange, COL_CHAR_CODE, "dataCharCode"
    
    SetDataName oFirstDataRange, COL_CHAR, "dataChar"
    SetDataName oFirstDataRange, COL_CHAR_OTHER_CASE, "dataOtherCase"
    SetDataName oFirstDataRange, COL_CHAR_AS_NAME_IS_OKAY, "dataNameOk"
    SetDataName oFirstDataRange, COL_CHAR_AT_START_IS_OKAY, "dataStartOk"
    SetDataName oFirstDataRange, COL_CHAR_AFTER_START_IS_OKAY, "dataAfterStartOk"
End Sub

Private Sub SetDataName(oFirstDataRange As Range, lCol As Long, sName As String)

    oFirstDataRange.Offset(, lCol - 1).Name = sName

End Sub

Public Sub StartDataGeneration()

    On Error GoTo Finally
    
    ' switiching off the following features gives
    ' a huge performance gain
    ActiveSheet.DisplayPageBreaks = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.Calculation = xlCalculationManual
    
    startDateTime = Now

    ' use this to only generate the ASCII characters
    'GenerateData 1, 255
    ' use this to only generate all unicode characters
    ' !ATTENTION! this may take several minutes and your
    '             PC may be unresponsive during this time !
    ' e.g. on an i7-2630QM CPU @ 2GHz it takes about 2 minutes
    '      until this sub is finished and then another 3
    '      minutes until Excel is responsive again
    '      these numbers are valid when you have an empty
    '      sheet without any calculations, etc.
    GenerateData

    Dbg "applying format"
    
    CondFormat
    
Finally:
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    Application.Calculation = xlCalculationAutomatic

    Dbg "finished in " & (Now - startDateTime)
End Sub

