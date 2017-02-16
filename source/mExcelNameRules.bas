Attribute VB_Name = "MExcelNameRules"
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

' These characters codes are invalid in all parts
' of an Excel Name (char only, at the start, after the start)
Private mAlwaysInvalid(1 To NAMES_NO_OF_INVALID_CHAR_CODES) As Long
' These are the only characters codes that are invalid at the
' start (not including the other invalid characters)
' Thus, if you want to know if a character is valid after the start
' * it must not exist in mInvalidAtStart
' * and of course it must also not exist in mInvalidAtStart
Private mInvalidAtStart(1 To NAMES_NO_OF_INVALID_CHAR_CODES_AT_START) As Long
' These characters are invalid name
' e.g. you cannot use them alone as a name: "c", "C", "r", "R"
Private mInvalidAsFullName(1 To 4) As Long

' These characters are invalid t the start of a name
' when you add the name to the workbook (worksheet is okay)
' these are only 2 characters
Private mInvalidAtStartOfWbOnly(1 To 2) As Long

' see Chr help: https://msdn.microsoft.com/en-us/library/office/gg264465.aspx
Public Const NAMES_MAX_UNICODE_CHARACTER_CODE = 65535

' see Microsoft Online Help: https://goo.gl/PsGUQj
Public Const NAMES_MAX_NAME_LEN = 255

' valid for Excel 2013, 2016
' see: Microsoft Help: https://goo.gl/VygIjJ
Private Const EXCEL_MAX_ROWS = 1048576
Private Const EXCEL_MAX_COLS = 16384

Private Sub Init()
    
    ' NOTE: we need to split the initialisation up to
    '       avoid "Procedure too large" errors
    Names_InitInvalidCharCodes1 mAlwaysInvalid
    
    Names_InitInvalidAtStart mInvalidAtStart
    
    mInvalidAsFullName(1) = 67 ' C
    mInvalidAsFullName(2) = 82 ' R
    mInvalidAsFullName(3) = 99 ' c
    mInvalidAsFullName(4) = 114 ' r
    
    mInvalidAtStartOfWbOnly(1) = 173
    mInvalidAtStartOfWbOnly(2) = 1600
    
End Sub

Private Sub LazyInit()
    If mInvalidAsFullName(1) = 0 Then Init
End Sub

' returns true if the value exists in the array
Private Function ArrayContains(arrSorted() As Long, lValueToSearch As Long) As Boolean

    Dim lTop As Long
    Dim lMiddle As Long
    Dim lBottom As Long
    
    lTop = UBound(arrSorted)
    lBottom = LBound(arrSorted)
    
    ArrayContains = False
    Do
        lMiddle = (lTop + lBottom) / 2
        If lValueToSearch > arrSorted(lMiddle) Then
           lBottom = lMiddle + 1
        ElseIf lValueToSearch = arrSorted(lMiddle) Then
            ArrayContains = True
            Exit Function
        Else
          lTop = lMiddle - 1
        End If
    Loop Until (lBottom > lTop)
End Function

Private Function AscWLong(sCharacter As String) As Long
    
    Dim lCharCode As Long
    
    If sCharacter = "" Then
        AscWLong = 0
        Exit Function
    End If
    
    lCharCode = AscW(sCharacter)
    ' note: AscW may return negative values
    ' see: https://msdn.microsoft.com/en-us/library/zew1e4wc(v=vs.90).aspx
    If lCharCode < 0 Then
        lCharCode = NAMES_MAX_UNICODE_CHARACTER_CODE + lCharCode + 1
    End If
    AscWLong = lCharCode
End Function

Public Function IsUcaseLetter(sCharToTest As String) As Boolean
    IsUcaseLetter = sCharToTest >= "A" And sCharToTest <= "Z"
End Function

Public Function IsDigit(sCharToTest As String) As Boolean
    IsDigit = sCharToTest >= "0" And sCharToTest <= "9"
End Function

' returns true when sInput starts with one or more letters,
' followed by at least one digit
' Examples
' true: "a1", "ab1", "a123", "abc123", "a1 ", "a1$"
' false: "", "a", "1", "a$1"
Private Function StartsWithLettersAndDigit(ByVal sInput As String) As Boolean

    If sInput = "" Then Exit Function
    
    sInput = UCase$(sInput)
    Dim sLetters As String
    sLetters = ""
    
    StartsWithLettersAndDigit = False
    
    Dim i As Integer
    Dim sCurChar As String
    For i = 1 To Len(sInput)
        sCurChar = Mid$(sInput, i, 1)
        If IsUcaseLetter(sCurChar) Then
            sLetters = sLetters & sCurChar
        ElseIf IsDigit(sCurChar) Then
            If sLetters = "" Then
                ' does not start with letters, but digits
                Exit Function
            Else
                ' starts with letter/s followed by at least on digit
                StartsWithLettersAndDigit = True
                Exit Function
            End If
        Else
            ' neither a letter nor a digit
            Exit Function
        End If
    Next i

End Function

Private Function LeftTrimToMaxLen(sInput As String, lMaxLen As Long) As String

    If Len(sInput) > lMaxLen Then
        LeftTrimToMaxLen = Left$(sInput, lMaxLen)
    Else
        LeftTrimToMaxLen = sInput
    End If

End Function

' Returns true, when input is a a backslash followed by
' exactly one character. Examples:
' * "\a" -> true
' * "\!" -> true
' * "\aa" -> false
' * "aa" -> false
Private Function IsSwitch(ByVal sInput As String) As Boolean
    IsSwitch = (Len(sInput) = 2) _
        And (Left$(sInput, 1) = "\")
End Function

Private Function ValidateName(ByRef sNameToTest As String, _
    Optional sReplaceChar As String = "") As Boolean

    ValidateName = False
    
    Dim bAdjustName As Boolean
    bAdjustName = sReplaceChar <> ""
    If bAdjustName Then
        ' since we will adjust the name it will always be valid
        ValidateName = True
        Debug.Assert Len(sReplaceChar) = 1
        Debug.Assert Names_IsCharValidAsName(sReplaceChar)
    End If
    
    If sNameToTest = "" Then
        If bAdjustName Then
            sNameToTest = sReplaceChar
        End If
        GoTo NameIsInvalid
    End If
    
    Dim lNameLen As Long
    lNameLen = Len(sNameToTest)
    
    ' max len is 255
    If lNameLen > NAMES_MAX_NAME_LEN Then
        If bAdjustName Then
            sNameToTest = LeftTrimToMaxLen(sNameToTest, NAMES_MAX_NAME_LEN)
        Else
            GoTo NameIsInvalid
        End If
    End If
    
    If lNameLen = 1 Then
        If Names_IsCharValidAsName(sNameToTest) Then GoTo NameIsValid
            
        If bAdjustName Then
            If Names_IsCharValidAfterStart(sNameToTest) Then
                ' note: just prepend the replace character to make it valid
                ' e.g. "?" and "?_" are invalid
                ' but "_?" is valid
                sNameToTest = sReplaceChar & sNameToTest
            Else
                ' this char is always invalid (e.g. "$", "!")
                sNameToTest = sReplaceChar
            End If
            GoTo NameIsValid
        Else
            GoTo NameIsInvalid
        End If
    Else
        Dim i As Integer
        Dim sCharToTest As String
        
        sCharToTest = Left$(sNameToTest, 1)
        ' test the first character
        Dim bInvalidStart As Boolean
        bInvalidStart = Not Names_IsCharValidAtStart(sCharToTest) _
            Or IsSwitch(sNameToTest)
        If bInvalidStart Then
            If bAdjustName Then
                If Names_IsCharValidAfterStart(sCharToTest) Then
                    ' start with the replace character which is for sure valid
                    ' e.g. "?x" -> "_?x"
                    sNameToTest = sReplaceChar & sNameToTest
                    ' note: we have prepended a character
                    lNameLen = lNameLen + 1
                Else
                    ' this character is never valid, so replace it
                    ' e.g. "$x" --> "_x"
                    sNameToTest = sReplaceChar & Right$(sNameToTest, Len(sNameToTest) - 1)
                End If
            Else
                GoTo NameIsInvalid
            End If
            Debug.Assert lNameLen = Len(sNameToTest)
        End If
        
        ' test the characters after the start
        Dim sAdjustedName As String
        sAdjustedName = Left$(sNameToTest, 1)
        For i = 2 To lNameLen
            sCharToTest = Mid$(sNameToTest, i, 1)
            If Names_IsCharValidAfterStart(sCharToTest) Then
                If bAdjustName Then
                    sAdjustedName = sAdjustedName & sCharToTest
                End If
            Else
                ' char is not valid after start
                If bAdjustName Then
                    sAdjustedName = sAdjustedName & sReplaceChar
                Else
                    GoTo NameIsInvalid
                End If
            End If
        Next i
        If bAdjustName Then
            sNameToTest = sAdjustedName
        End If
        
        ' this covers reference like names e.g. "R1C1", "A1", ..
        If StartsWithLettersAndDigit(sNameToTest) Then
            If bAdjustName Then
                ' just prepend the replace character: "_R1C1" is valid
                sNameToTest = sReplaceChar & sNameToTest
            Else
                GoTo NameIsInvalid
            End If
        End If
        
    End If
    GoTo NameIsValid
    
NameIsValid:
    ValidateName = True
    If bAdjustName Then
        ' note: we may have prepended the replace char, so we
        '       must check the length again
        sNameToTest = LeftTrimToMaxLen(sNameToTest, NAMES_MAX_NAME_LEN)
    End If
    Exit Function
    
NameIsInvalid:
    ValidateName = False
End Function

' returns true if sNameToTest is a valid Excel Name
Public Function Names_IsValidName(sNameToTest As String) As Boolean

    Names_IsValidName = ValidateName(sNameToTest)
    
End Function

' will return an adjusted version of the name that only contains
' valid characters
' the function may:
' * replace a single character
'   "0" -> "_"
'   "$" -> "_"
' * prepend the replace character to avoid invalid names
'   "c" -> "_c"
'   "?" -> "_?"
'   "A1" -> "_A1"
' * replace invalid characters
'   "ab$" -> "ab_"
' * trim the string if it is too long
'
' notes:
' * this function may produce the same output for
'   different input string, so you must handle duplicates
'   e.g. "0" and "$" will both return "_"
' * also remember that names in Excel are case-insensitive
'   e.g. "Abc" and "aBC" is the same name
Public Function Names_AdjustName(sNameToTest As String, _
    Optional sReplaceChar As String = "_") As String

    If Not Names_IsCharValidAtStart(sReplaceChar) Then
        sReplaceChar = "_"
    End If
    
    Dim sResult As String
    sResult = sNameToTest
    ValidateName sResult, sReplaceChar
    Names_AdjustName = sResult
    
End Function


' This function returns true when the character can be used after
' the start of a name.
' e.g. numbers (0-9) cannot be used as name alone and must also
'      not be used at the start but may be okay after the start
' note: when this function returns true, it does not yet mean, that
'       you can always build valid names:
'       Example: IsCharValidAfterStart("1") returns True
'                * the name "A1xx" is valid
'                * the name "A1" is not valid (cell-reference)
Public Function Names_IsCharValidAfterStart(sCharacter As String) As Boolean
    Names_IsCharValidAfterStart = Names_IsCharCodeValidAfterStart(AscWLong(sCharacter))
End Function

Public Function Names_IsCharCodeValidAfterStart(lCharCode As Long) As Boolean
    LazyInit
    Names_IsCharCodeValidAfterStart = Not ArrayContains(mAlwaysInvalid, lCharCode)
End Function

' This function returns true when the character can be used at
' the start of a name (but not necesarily alone as a full name)
' e.g. IsCharValidAtStart("C") returns true
'      * the name "AC" is valid
'      * the name "AC23" is invalid (cell reference)
'      * the name "C" (character alone as full name) is invalid
' Note: when a character is valid the start, it also means
'       that it is valid after the start
Public Function Names_IsCharValidAtStart(sCharacter As String) As Boolean
    Names_IsCharValidAtStart = Names_IsCharCodeValidAtStart(AscWLong(sCharacter))
End Function

Public Function Names_IsCharCodeValidAtStart(lCharCode As Long) As Boolean
    LazyInit
    Names_IsCharCodeValidAtStart = Names_IsCharCodeValidAfterStart(lCharCode) _
        And (Not ArrayContains(mInvalidAtStart, lCharCode)) _
        And (Not ArrayContains(mInvalidAtStartOfWbOnly, lCharCode))
End Function

' This function returns true when the character can be used as
' a full name
' Examples:
' * IsCharValidAtStart("A") returns true
' * IsCharValidAtStart("C") returns false
' * IsCharValidAtStart("$") returns false ("$" is always invalid)
Public Function Names_IsCharValidAsName(sCharacter As String) As Boolean
    Names_IsCharValidAsName = Names_IsCharCodeValidAsName(AscWLong(sCharacter))
End Function

Public Function Names_IsCharCodeValidAsName(lCharCode As Long) As Boolean
    LazyInit
    Names_IsCharCodeValidAsName = Names_IsCharCodeValidAtStart(lCharCode) _
        And (Not ArrayContains(mInvalidAsFullName, lCharCode))
End Function

