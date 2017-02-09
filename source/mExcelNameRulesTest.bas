Attribute VB_Name = "mExcelNameRulesTest"
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

Private Const ALWAYS_VALID_UNICODE = 24475 ' some asian character

' This function will loop over all data in the sheet
' with the generated data and call the check-functions
' and compare if the VBA functions return the same value.
Public Sub Names_Tests_CompareSheetDataToCompuation()
    Dim oRangeCharCodes As Range
    Set oRangeCharCodes = ActiveSheet.Range("dataCharCode")
    
    Dim oRangeValidAsName As Range
    Set oRangeValidAsName = ActiveSheet.Range("dataNameOk")
    
    Dim oRangeStartOk As Range
    Set oRangeStartOk = ActiveSheet.Range("dataStartOk")
    
    Dim oRangeAfterStartOk As Range
    Set oRangeAfterStartOk = ActiveSheet.Range("dataAfterStartOk")
    
    Debug.Print "starting test"
    Dim oCell As Variant
    Dim lErrorCount As Long
    lErrorCount = 0
    Dim lIdx As Long
    For lIdx = 1 To NAMES_MAX_UNICODE_CHARACTER_CODE
        Dim lCharCode As Long
        Dim bIsValidCalc As Boolean
        Dim bIsValidInSheet As Boolean
        
        lCharCode = oRangeCharCodes.Cells(lIdx)
        
        ' check if char as full-name is okay
        bIsValidCalc = Names_IsCharCodeValidAsName(lCharCode)
        bIsValidInSheet = oRangeValidAsName.Cells(lIdx).Value
        If bIsValidInSheet <> bIsValidCalc Then
            lErrorCount = lErrorCount + 1
            Debug.Print "Mismatch: FullName! " & lCharCode & "=" & bIsValidCalc _
                & " sheet: " & bIsValidInSheet
        End If
        
        ' check if char at start is is okay
        bIsValidCalc = Names_IsCharCodeValidAtStart(lCharCode)
        bIsValidInSheet = oRangeStartOk.Cells(lIdx).Value
        If bIsValidInSheet <> bIsValidCalc Then
            lErrorCount = lErrorCount + 1
            Debug.Print "Mismatch: AtStart! " & lCharCode & "=" & bIsValidCalc _
                & " sheet: " & bIsValidInSheet
        End If
    
        ' check if char after start is is okay
        bIsValidCalc = Names_IsCharCodeValidAfterStart(lCharCode)
        bIsValidInSheet = oRangeAfterStartOk.Cells(lIdx).Value
        If bIsValidInSheet <> bIsValidCalc Then
            lErrorCount = lErrorCount + 1
            Debug.Print "Mismatch: AfterStart! " & lCharCode & "=" & bIsValidCalc _
                & " sheet: " & bIsValidInSheet
        End If
    
    Next lIdx
    
    If lErrorCount = 0 Then
        Debug.Print "Test has finished successfully"
    Else
        Debug.Print "Test has FAILED with " & lErrorCount & " error/s"
    End If
    
End Sub

Private Sub AssertNamesValid(ParamArray varrValidNames() As Variant)
    
    Dim vName As Variant
    For Each vName In varrValidNames
        Debug.Assert Names_IsValidName("" & vName)
    Next
End Sub

Private Sub AssertNamesInvalid(ParamArray varrInvalidNames() As Variant)
    
    Dim vName As Variant
    For Each vName In varrInvalidNames
        Debug.Assert Not Names_IsValidName("" & vName)
    Next
End Sub

Private Sub Test_IsValidName_SingleCharacter()
    ' c, r: not valid as full-name
    AssertNamesInvalid "c", "C", "r", "R"
    
    ' $ is never valid
    AssertNamesInvalid "$", "!"
    ' e.g. "1" is not valid as full name
    ' but valid after start e.g. "_1"
    AssertNamesInvalid "1", "2"
    AssertNamesValid "_1", "_2"
    ' also "A1", etc. are not valid: cell-references
    AssertNamesInvalid "A1", "A2xx"
    
    AssertNamesValid "A", "Z"
    
    ' test some VALID characters
    AssertNamesValid "\", "_", ChrW$(ALWAYS_VALID_UNICODE)
End Sub

Private Sub Test_IsValidName_MultipleCharacters()
    ' $ is never valid
    AssertNamesInvalid "$A"
    ' e.g. "1" is not valid at the start
    AssertNamesInvalid "1A"
    
    ' test some VALID characters
    AssertNamesValid "AA", "ZA", "\A", "_A", ChrW$(ALWAYS_VALID_UNICODE) & "A"
    
    ' c, r: are valid at the start
    AssertNamesValid "cA", "CA", "rA", "RA"
    
    ' e.g. "c" is valid after start (but not at start)
    AssertNamesValid "rC", "AC"
End Sub

Private Sub Test_IsValidName_Space()
    AssertNamesInvalid " ", " A", "A ", " A "
End Sub

Private Sub Test_IsValidName_CellReferences()
    AssertNamesInvalid "A1", "ABC1", "A123", "Z100", "Z$100", "A1:B2"
    AssertNamesInvalid "R1C1", "R10", "C10", "R123C123"
End Sub

Public Sub Names_Tests_IsValidName()
    
    Const TEST_NAME = "Names_Tests_IsValidName"
    
    Debug.Print "starting test: " & TEST_NAME
    
    ' empty string is not valid
    Debug.Assert Not Names_IsValidName("")
    
    Test_IsValidName_SingleCharacter
    Test_IsValidName_MultipleCharacters
    Test_IsValidName_Space
    Test_IsValidName_CellReferences
    
    ' test identifier that is too long
    Dim sNameTooLong As String
    sNameTooLong = String(256, "_")
    Debug.Assert Not Names_IsValidName(sNameTooLong)
    
    Debug.Print "finished test: " & TEST_NAME
    
End Sub

Private Sub AssertAdjustName(sInput As String, sExpected As String, _
    Optional sReplaceChar As String = "_")
    
    Dim sAdjustedName As String
    sAdjustedName = Names_AdjustName(sInput, sReplaceChar)
    Debug.Assert sAdjustedName = sExpected
    
End Sub

Private Sub AssertAdjustName4ValidNames(ParamArray varrValidNames() As Variant)

    Dim vName As Variant
    Dim sName As String
    For Each vName In varrValidNames
        sName = "" & vName
        AssertAdjustName sName, sName
    Next
    ' again with other replacement character
    For Each vName In varrValidNames
        sName = "" & vName
        AssertAdjustName sName, sName, "\"
    Next
End Sub

Private Sub Names_Tests_AdjustLongName(ByVal sChar As String)

    Dim sExpected As String
    sExpected = String(NAMES_MAX_NAME_LEN, sChar)
    
    Dim sTooLong As String
    ' note: we can even add invalid characters
    sTooLong = String(NAMES_MAX_NAME_LEN, sChar) & "$!"
    
    Dim sAdjusted As String
    sAdjusted = Names_AdjustName(sTooLong)
    Debug.Assert Len(sAdjusted) = NAMES_MAX_NAME_LEN
    Debug.Assert sAdjusted = sExpected
End Sub

' asserts that each name will be prepended with the repalace char
' and no other replacements were made
Private Sub AssertAdjustNamePrependsReplacement(ParamArray varrInvalidNames() As Variant)

    Dim vName As Variant
    Dim sName As String
    Dim sExpected As String
    For Each vName In varrInvalidNames
        sName = "" & vName
        sExpected = "_" & sName
        AssertAdjustName sName, sExpected
    Next
    ' again with other replacement character
    For Each vName In varrInvalidNames
        sName = "" & vName
        sExpected = "\" & sName
        AssertAdjustName sName, sExpected, "\"
    Next
End Sub

Public Sub Names_Tests_AdjustName()
    
    Const TEST_NAME = "Names_Tests_IsValidName"
    
    Debug.Print "starting test: " & TEST_NAME
    
    AssertAdjustName "", "_"
    ' single (always) invalid char  will return the replace char
    AssertAdjustName "$", "_"
    AssertAdjustName "!", "_"
    AssertAdjustName "$", "\", "\"
    AssertAdjustName "!", "\", "\"
    ' single invalid char at start will prepend the replace char
    AssertAdjustName "c", "_c"
    AssertAdjustName "c", "\c", "\"
    AssertAdjustName "1", "_1"
    AssertAdjustName "1", "\1", "\"
    
    ' spaces
    AssertAdjustName " ", "_"
    AssertAdjustName " c", "_c"
    AssertAdjustName " c ", "_c_"
    AssertAdjustName " x", "_x"
    AssertAdjustName " x ", "_x_"
    AssertAdjustName " x ", "_x_"
    
    ' invalid start char but valid afterwards
    AssertAdjustName "?x", "_?x"
    ' (always) invalid start char
    AssertAdjustName "$x", "_x"
    
    ' invalid cell-ref
    AssertAdjustNamePrependsReplacement "A1", "a1"
    ' R1C1 like references
    AssertAdjustNamePrependsReplacement "R1048576C1" _
        , "R1048577C1", "R1C16384", "R1C16385", "R1C", "R1D" _
        , "R1C16385xxx", "A1", "XFD1048577", "XFD1048578", "XFE1048577"
    
    ' too long
    Names_Tests_AdjustLongName "x"
    Names_Tests_AdjustLongName ChrW$(ALWAYS_VALID_UNICODE)
    
    ' VALID NAMES will not be changed
    ' single char
    AssertAdjustName4ValidNames "A", "a", "Z", "z"
    ' single char unicode
    AssertAdjustName4ValidNames ChrW$(ALWAYS_VALID_UNICODE)
    ' multi char
    AssertAdjustName4ValidNames "abc", "cd"
    ' multi char unicode
    AssertAdjustName4ValidNames "ab" & ChrW$(ALWAYS_VALID_UNICODE), _
        ChrW$(ALWAYS_VALID_UNICODE) & "ab", _
        "ab" & ChrW$(ALWAYS_VALID_UNICODE) & "xy"
        
    Debug.Print "finished test: " & TEST_NAME
    
End Sub
