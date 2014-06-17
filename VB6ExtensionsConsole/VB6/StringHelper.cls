VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StringHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Const ERR_FORMAT_EXCEPTION As Long = vbObjectError Or 9001
Private Const ERR_SOURCE As String = "StringHelper"
Private Const ERR_MSG_INVALID_FORMAT_STRING As String = "Invalid format string."
Private Const ERR_MSG_FORMAT_EXCEPTION As String = "The number indicating an argument to format is less than zero, or greater than or equal to the length of the args array."

Private Type tString
    PaddingCharacter As String * 1
    EscapeSequences As New Collection
    NumericSpecifiers As New Collection
    DateTimeSpecifiers As New Collection
End Type

Private Const PADDING_CHAR As String * 1 = " "

Private this As tString
Option Base 0
Option Explicit

Private Sub Class_Initialize()
    
    If this.PaddingCharacter = vbNullString Then this.PaddingCharacter = PADDING_CHAR
    
    InitEscapeSequences
    InitNumericSpecifiers
    InitDateTimeSpecifiers
    
End Sub

Private Sub InitEscapeSequences()
    
    Dim factory As New EscapeSequence
    Set this.EscapeSequences = New Collection
    
    this.EscapeSequences.Add factory.Create("\n", vbNewLine)
    this.EscapeSequences.Add factory.Create("\q", Chr$(34))
    this.EscapeSequences.Add factory.Create("\t", vbTab)
    this.EscapeSequences.Add factory.Create("\a", Chr$(7))
    this.EscapeSequences.Add factory.Create("\b", Chr$(8))
    this.EscapeSequences.Add factory.Create("\v", Chr$(13))
    this.EscapeSequences.Add factory.Create("\f", Chr$(14))
    this.EscapeSequences.Add factory.Create("\r", Chr$(15))
    this.EscapeSequences.Add factory.Create("\\x(\w{2})", 0, True, Hexadecimal)
    this.EscapeSequences.Add factory.Create("\\(\d{3})", 0, True, Octal)
    
    Set factory = Nothing
    
End Sub

Private Sub InitNumericSpecifiers()
    
    Set this.NumericSpecifiers = New Collection
    this.NumericSpecifiers.Add New CurrencyStringFormatIdentifier
    this.NumericSpecifiers.Add New DecimalStringFormatIdentifier
    this.NumericSpecifiers.Add New GeneralNumericStringFormatIdentifier
    this.NumericSpecifiers.Add New PercentStringFormatIdentifier
    this.NumericSpecifiers.Add New NumericStringFormatIdentifier
    this.NumericSpecifiers.Add New FixedPointStringFormatIdentifier
    this.NumericSpecifiers.Add New ExponentialStringFormatIdentifier
    this.NumericSpecifiers.Add New HexStringFormatIdentifier
    this.NumericSpecifiers.Add New RoundTripStringFormatIdentifier
    this.NumericSpecifiers.Add New NumericPaddingStringFormatIdentifier
    
End Sub

Private Sub InitDateTimeSpecifiers()
    
    Set this.DateTimeSpecifiers = New Collection
    this.DateTimeSpecifiers.Add New CustomDateFormatIdentifier
    this.DateTimeSpecifiers.Add New FullDateLongStringFormatSpecifier
    this.DateTimeSpecifiers.Add New FullDateShortStringFormatIdentifier
    this.DateTimeSpecifiers.Add New GeneralLongDateTimeStringFormatIdentifier
    this.DateTimeSpecifiers.Add New GeneralShortDateTimeStringFormatIdentifier
    this.DateTimeSpecifiers.Add New LongDateFormatIdentifier
    this.DateTimeSpecifiers.Add New LongTimeStringFormatIdentifier
    this.DateTimeSpecifiers.Add New ShortDateFormatIdentifier
    this.DateTimeSpecifiers.Add New SortableDateTimeStringFormatIdentifier
    
End Sub

Public Function StringFormat(format_string As String, values() As Variant) As String
    
    Dim result As String
    result = format_string
    
    Dim specifiers As Collection
    Dim specifier As StringFormatSpecifier
    Set specifiers = GetFormatSpecifiers(result, UBound(values) + 1)
    
    Dim useLiteral As Boolean 'when format_string starts with "@", escapes are not replaced (string is treated as a literal string with placeholders)
    useLiteral = StringStartsWith("@", result)
    If useLiteral Then result = Right(result, Len(result) - 1) 'remove the "@" literal specifier from the result string
    
    'replace escaped backslashes with 'ESC' character [Chr$(27)] to optimize escape sequences evaluation:
    If Not useLiteral And StringContains(result, "\\") Then result = Replace(result, "\\", Chr$(27))
    
    Dim formattedValue As String
    Dim alignmentPadding As Integer
    Dim Identifier As IStringFormatIdentifier
    Dim identifierFound As Boolean
    
    For Each specifier In specifiers
        
        formattedValue = values(specifier.Index)
        identifierFound = (specifier.Identifier = vbNullString)
        
        If IsNumeric(values(specifier.Index)) Then
            
            For Each Identifier In this.NumericSpecifiers
                If Identifier.IsIdentifierMatch(specifier) Then
                    
                    identifierFound = True
                    formattedValue = Identifier.GetFormattedValue(values(specifier.Index), specifier)
                    Exit For
                    
                End If
            Next
            
        ElseIf TypeName(values(specifier.Index)) = "Date" Then
            
            For Each Identifier In this.DateTimeSpecifiers
                If Identifier.IsIdentifierMatch(specifier) Then
                    
                    identifierFound = True
                    formattedValue = Identifier.GetFormattedValue(values(specifier.Index), specifier)
                    Exit For
                    
                End If
            Next
            
        End If
        
        If Not identifierFound And Not IsNumeric(specifier.Identifier) Then Err.Raise ERR_FORMAT_EXCEPTION, ERR_SOURCE, ERR_MSG_INVALID_FORMAT_STRING
        ApplyAlignmentSpecifier formattedValue, specifier
                
        'replace all occurrences of placeholder {i} with their formatted values:
        result = Replace(result, specifier.ToString, formattedValue)
        
    Next
    
    Dim escape As EscapeSequence
    If Not useLiteral And StringContains(result, "\") Then
        For Each escape In this.EscapeSequences
            escape.Execute result
        Next
    End If
    
    If Not useLiteral And StringContains(result, Chr$(27)) Then result = Replace(result, Chr$(27), "\")
    StringFormat = result
    
End Function

Private Function GetFormatSpecifiers(ByVal format_string As String, valuesCount As Integer) As Collection
'executes a regular expression against format_string to extract all placeholders into a MatchCollection
'note: THIS is the performance bottleneck of StringFormat!

    Dim regex As New RegExp
    Dim matches As MatchCollection
    Dim thisMatch As match
    
    Dim result As New Collection
    Dim specifier As StringFormatSpecifier
    
    Dim csvIndices As String
    Dim uniqueCount As Integer
    Dim largestIndex As Integer
    
    regex.pattern = "\{(\w+)(\,\-?\d+)?(\:[^}]+)?\}"
    
    ' literal {
    ' [1] numbered captured group, any number of repetitions (Index)
    '    alphanumeric, one or more repetitions
    ' [2] numbered captured group, zero or one repetitions (AlignmentSpecifier)
    '    literal ,
    '    literal -, zero or one repetitions
    '    any digit, one or more repetitions
    ' [3] numbered captured group, zero or one repetitions (FormatSpecifier)
    '    literal :
    '    any character except '}', one or more repetitions
    ' literal }
    
    regex.IgnoreCase = True
    regex.Global = True
    
    Set matches = regex.Execute(format_string)
    For Each thisMatch In matches
        
        Set specifier = New StringFormatSpecifier
        specifier.Index = CInt(thisMatch.SubMatches(0))
        
        If Not StringContains(csvIndices, specifier.Index & ",") Then
            uniqueCount = uniqueCount + 1
            csvIndices = csvIndices & specifier.Index & ","
        End If
        If specifier.Index > largestIndex Then largestIndex = specifier.Index
        
        If Not thisMatch.SubMatches(1) = vbEmpty Then specifier.Alignment = CInt(Replace(CStr(thisMatch.SubMatches(1)), ",", vbNullString))
        If Not thisMatch.SubMatches(2) = vbEmpty Then
            specifier.Identifier = Left(Replace(CStr(thisMatch.SubMatches(2)), ":", vbNullString), 1)
            specifier.CustomSpecifier = Replace(CStr(thisMatch.SubMatches(2)), ":" & specifier.Identifier, vbNullString)
        End If
        
        result.Add specifier
    Next
    
    If matches.Count > 0 And ((uniqueCount <> valuesCount) Or (largestIndex >= uniqueCount) Or valuesCount = 0) Then Err.Raise ERR_FORMAT_EXCEPTION, ERR_SOURCE, ERR_MSG_FORMAT_EXCEPTION
    
    Set GetFormatSpecifiers = result
    Set regex = Nothing
    Set matches = Nothing
    
End Function

Private Sub ApplyAlignmentSpecifier(ByRef formattedValue As String, specifier As StringFormatSpecifier)
    
    Dim alignmentPadding As Integer
    alignmentPadding = Abs(specifier.Alignment)
    
    If specifier.Alignment < 0 Then
        
        'negative: left-justified alignment
        If alignmentPadding - Len(formattedValue) > 0 Then _
            formattedValue = formattedValue & String$(alignmentPadding - Len(formattedValue), this.PaddingCharacter)
    
    ElseIf specifier.Alignment > 0 Then
        
        'positive: right-justified alignment
        If alignmentPadding - Len(formattedValue) > 0 Then _
            formattedValue = String$(alignmentPadding - Len(formattedValue), this.PaddingCharacter) & formattedValue
            
    End If
End Sub

Public Sub TestMethod_AllSpecifiers()
    
    Dim integerValue As Integer
    Dim numericValue As Double
    Dim dateTimeValue As Date
    Dim Precision As Integer
    Dim Alignment As Integer
        
    integerValue = 123
    numericValue = 123 / 456
    dateTimeValue = Now
    
    this.PaddingCharacter = "."
    
    Debug.Print
    Debug.Print "IntegerValue : " & integerValue
    Debug.Print "NumericValue : " & numericValue
    Debug.Print "DateTimeValue: " & dateTimeValue
    Debug.Print "Padding      : Chr$(" & Asc(this.PaddingCharacter) & ") / """ & this.PaddingCharacter & """"
    Debug.Print
    
    Dim specifier As IStringFormatIdentifier
    Dim escape As EscapeSequence
    
    Debug.Print "EscapeSequences"
    For Each escape In this.EscapeSequences
        If Not escape.IsAsciiCharacter Then
            Debug.Print StringHelperModule.StringFormat("@{0:20}", escape.EscapeString) & StringHelperModule.StringFormat("\t{0:20} (ascii {1})", escape.ReplacementString, Asc(escape.ReplacementString))
        ElseIf escape.AsciiBase = Octal Then
            Debug.Print StringHelperModule.StringFormat("\" & Format(Oct(65), "000") & " (ascii: 65)")
        ElseIf escape.AsciiBase = Hexadecimal Then
            Debug.Print StringHelperModule.StringFormat("\x" & Replace(CStr(Hex(65)), "&H", "") & " (ascii: 65)")
        End If
    Next
    Debug.Print
    
    Debug.Print "DateTimeSpecifiers:"
    For Each specifier In this.DateTimeSpecifiers
        Debug.Print specifier.Identifier & StringHelperModule.StringFormat("\t{0,48:" & specifier.Identifier & "}", dateTimeValue)
    Next
    Debug.Print
        
    Alignment = 20
    For Precision = 0 To 5
        Debug.Print
        Debug.Print "Precision    : " & Precision
        Debug.Print "Alignment    : " & Alignment
        Debug.Print
        Debug.Print "NumericSpecifiers:"
        
        For Each specifier In this.NumericSpecifiers
            If specifier.Identifier <> "0" Then
                Debug.Print UCase(specifier.Identifier) & StringHelperModule.StringFormat("\t{0," & Alignment & ":" & UCase(specifier.Identifier) & Precision & "}\t\t{1," & Alignment & ":" & UCase(specifier.Identifier) & Precision & "}", integerValue, numericValue)
                Debug.Print LCase(specifier.Identifier) & StringHelperModule.StringFormat("\t{0," & Alignment & ":" & LCase(specifier.Identifier) & Precision & "}\t\t{1," & Alignment & ":" & LCase(specifier.Identifier) & Precision & "}", integerValue, numericValue)
            End If
        Next
                
    Next
    
    
    this.PaddingCharacter = PADDING_CHAR
    
End Sub

'Renvoie TRUE si string_source contient find_text (case insensitive by default).
Public Function StringContains(ByVal string_source As String, ByVal find_text As String, Optional ByVal caseSensitive As Boolean = False) As Boolean
        
    Dim compareMethod As VbCompareMethod
    
    If caseSensitive Then
        compareMethod = vbBinaryCompare
    Else
        compareMethod = vbTextCompare
    End If
    
    StringContains = (InStr(1, string_source, find_text, compareMethod) <> 0)
    
End Function

'Renvoie TRUE si string_source contient n'importe laquelle des valeurs specifiees.
Public Function StringContainsAny(ByVal string_source As String, ByVal caseSensitive As Boolean, find_strings() As Variant) As Boolean
    
    Dim find As String, i As Integer, found As Boolean
    
    For i = LBound(find_strings) To UBound(find_strings)
        
        find = CStr(find_strings(i))
        found = StringContains(string_source, find, caseSensitive)
        
        If found Then Exit For
    Next
    
    StringContainsAny = found
    
End Function

'Renvoie TRUE si string_source égale n'importe laquelle des valeurs specifiées.
Public Function StringMatchesAny(ByVal string_source As String, find_strings() As Variant) As Boolean
    
    Dim find As String, src As String, i As Integer, found As Boolean
    
    For i = LBound(find_strings) To UBound(find_strings)
        
        find = CStr(find_strings(i))
        found = (string_source = find)
        
        If found Then Exit For
    Next
    
    StringMatchesAny = found
    
End Function

'Renvoie TRUE si string_source égale toutes les valeurs specifiées.
Public Function StringMatchesAll(ByVal string_source As String, find_strings() As Variant) As Boolean
    
    Dim find As String, i As Integer, match As Boolean
    
    For i = LBound(find_strings) To UBound(find_strings)
        
        find = CStr(find_strings(i))
        match = (string_source = find)
        
        If Not match Then Exit For
    Next
    
    StringMatchesAll = match
    
End Function

'Renvoie TRUE si string_source débute par find_text (case sensitive).
Public Function StringStartsWith(ByVal find_text As String, ByVal string_source As String, Optional ByVal caseSensitive As Boolean = True) As Boolean
    
    If Not caseSensitive Then
        string_source = LCase$(string_source)
        find_text = LCase$(find_text)
    End If

    StringStartsWith = (Left$(string_source, LenB(find_text) / 2) = find_text)

End Function

Public Function StringStartsWithAny(ByVal string_source As String, ByVal caseSensitive As Boolean, find_strings() As Variant) As Boolean
    
    Dim find As String, i As Integer, found As Boolean
    
    For i = LBound(find_strings) To UBound(find_strings)
        
        find = CStr(find_strings(i))
        found = StringStartsWith(find, string_source, caseSensitive)
        
        If found Then Exit For
    Next
    
    StringStartsWithAny = found
    
End Function

'Renvoie TRUE si string_source se termine par find_text (case sensitive).
Public Function StringEndsWith(ByVal find_text As String, ByVal string_source As String, Optional ByVal caseSensitive As Boolean = True) As Boolean
        
    If Not caseSensitive Then
        string_source = LCase$(string_source)
        find_text = LCase$(find_text)
    End If
    
    StringEndsWith = (Right$(string_source, LenB(find_text) / 2) = find_text)

End Function

Public Function StringEndsWithAny(ByVal string_source As String, ByVal caseSensitive As Boolean, find_strings() As Variant) As Boolean
    
    Dim find As String, i As Integer, found As Boolean
    
    For i = LBound(find_strings) To UBound(find_strings)
        
        find = CStr(find_strings(i))
        found = StringEndsWith(find, string_source, caseSensitive)
        
        If found Then Exit For
    Next
    
    StringEndsWithAny = found
    
End Function


