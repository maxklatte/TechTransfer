Attribute VB_Name = "StringHelpers"
' Module: StringHelpers
' Shared string helper functions for formatting, cleanup, display

Option Explicit

' --- Significant Digits ---
Public Function SigStr(ByVal num As Double, Optional ByVal sigDigits As Integer = 3) As String
    Dim exponent As Integer
    Dim scaleFactor As Double
    Dim rounded As Double
    Dim decimalPlaces As Integer
    Dim fmt As String
    Dim result As String

    If num = 0 Then
        result = "0." & String(sigDigits - 1, "0")
        SigStr = result
        Exit Function
    End If

    exponent = Int(Log(Abs(num)) / Log(10))
    scaleFactor = 10 ^ (sigDigits - exponent - 1)
    rounded = Round(num * scaleFactor) / scaleFactor

    decimalPlaces = sigDigits - Int(Log(Abs(rounded)) / Log(10)) - 1
    If decimalPlaces < 0 Then decimalPlaces = 0

    If decimalPlaces = 0 Then
        fmt = "0"
    Else
        fmt = "0." & String(decimalPlaces, "0")
    End If

    result = format$(rounded, fmt)
    result = Replace(result, Application.International(wdDecimalSeparator), ".")
    
    SigStr = result
End Function


' --- Padding ---
Public Function PadLeft(text As String, width As Integer, Optional char As String = " ") As String
    PadLeft = String(width - Len(text), char) & text
End Function

Public Function PadRight(text As String, width As Integer, Optional char As String = " ") As String
    PadRight = text & String(width - Len(text), char)
End Function

' --- Case Formatting ---
Public Function CapFirst(text As String) As String
    If Len(text) = 0 Then
        CapFirst = ""
    Else
        CapFirst = UCase(Left(text, 1)) & LCase(Mid(text, 2))
    End If
End Function

Public Function ToTitleCase(text As String) As String
    Dim parts() As String, i As Integer
    parts = Split(LCase(text))
    For i = 0 To UBound(parts)
        If Len(parts(i)) > 0 Then parts(i) = UCase(Left(parts(i), 1)) & Mid(parts(i), 2)
    Next i
    ToTitleCase = Join(parts, " ")
End Function

' --- Cleaning ---
Public Function StripNonAlpha(text As String) As String
    Dim i As Long, c As String, result As String
    result = ""
    For i = 1 To Len(text)
        c = Mid(text, i, 1)
        If c Like "[A-Za-z]" Then result = result & c
    Next i
    StripNonAlpha = result
End Function

Public Function OnlyDigits(text As String) As String
    Dim i As Long, c As String, result As String
    result = ""
    For i = 1 To Len(text)
        c = Mid(text, i, 1)
        If c Like "#" Then result = result & c
    Next i
    OnlyDigits = result
End Function

Public Function StripWhitespace(text As String) As String
    Dim i As Long, c As String, result As String
    result = ""
    For i = 1 To Len(text)
        c = Mid(text, i, 1)
        If c <> " " And c <> vbTab And c <> vbCr And c <> vbLf Then
            result = result & c
        End If
    Next i
    StripWhitespace = result
End Function



' --- Utilities ---
Public Function TruncateStr(text As String, maxLength As Integer) As String
    If Len(text) <= maxLength Then
        TruncateStr = text
    Else
        TruncateStr = Left(text, maxLength - 3) & "..."
    End If
End Function

Public Function IsNullOrEmpty(text As String) As Boolean
    IsNullOrEmpty = (Len(Trim(text)) = 0)
End Function

Public Function Pluralize(word As String, count As Long) As String
    If count = 1 Then
        Pluralize = word
    Else
        Pluralize = word & "s"
    End If
End Function

Public Function SanitizeDisplayString(ByVal raw As String) As String
    Dim cleaned As String
    Dim i As Integer, c As String, ascii As Long
    cleaned = ""

    For i = 1 To Len(raw)
        c = Mid(raw, i, 1)
        ascii = AscW(c)

        ' Allow printable ASCII + common symbols
        If (ascii >= 32 And ascii <= 126) Or _
           ascii = 160 Or ascii = 169 Or ascii = 174 Or ascii = 8482 Then
            Select Case ascii
                Case 160: cleaned = cleaned & " "
                Case 8482: cleaned = cleaned & "[tm]"
                Case 174: cleaned = cleaned & "[R]"
                Case 169: cleaned = cleaned & "[C]"
                Case Else: cleaned = cleaned & c
            End Select
        ElseIf ascii = 10 Or ascii = 13 Then
            cleaned = cleaned & " "
        End If
    Next i

    SanitizeDisplayString = cleaned
End Function

'=== MODULE: StringHelpers - XML strings ===

' Sanitize string for safe use as XML attribute or text content
Public Function SanitizeXmlValue(ByVal raw As String) As String
    Dim result As String
    Dim i As Long, c As String, code As Long
    result = ""

    ' Filter invalid XML 1.0 chars
    For i = 1 To Len(raw)
        c = Mid(raw, i, 1)
        code = AscW(c)

        If (code = 9 Or code = 10 Or code = 13) Or _
           (code >= 32 And code <= 55295) Or _
           (code >= 57344 And code <= 65533) Then
            result = result & c
        End If
    Next i

    '  — No more Replace("&", "&amp;"), Replace("<","&lt;"), etc. —
    'result = Replace(result, "&", "&amp;")
    'result = Replace(result, "<", "&lt;")
    'result = Replace(result, ">", "&gt;")
    'result = Replace(result, """", "&quot;")
    'result = Replace(result, "'", "&apos;")

    SanitizeXmlValue = result
End Function
Public Function UnescapeXmlValue(ByVal xmlEscaped As String) As String
    Dim tmp As String
    tmp = xmlEscaped

    '— Always un-escape the ampersand first, to handle cases like “&amp;lt;”
    tmp = Replace(tmp, "&amp;", "&")

    '— Then turn the other five XML entities into real characters
    tmp = Replace(tmp, "&lt;", "<")
    tmp = Replace(tmp, "&gt;", ">")
    tmp = Replace(tmp, "&quot;", """")
    tmp = Replace(tmp, "&apos;", "'")

    UnescapeXmlValue = tmp
End Function

Public Function IsWellFormedXml(xmlString As String) As Boolean
    Dim xmlDoc As Object
    Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
    xmlDoc.async = False
    xmlDoc.validateOnParse = False
    xmlDoc.resolveExternals = False

    IsWellFormedXml = xmlDoc.LoadXML(xmlString)
End Function

