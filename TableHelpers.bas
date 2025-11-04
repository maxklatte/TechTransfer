Attribute VB_Name = "TableHelpers"
' === Module: RowValidationHelpers ===
Option Explicit

' === Validates that a row is suitable for inserting a clone below it ===
Public Function IsValidRowForInsertion(r As row, Optional expectedRegionTitle As String = "") As Boolean
    On Error GoTo ErrorHandler

    If r Is Nothing Then
        Debug.Print "[IsValidRow] Row is Nothing."
        IsValidRowForInsertion = False
        Exit Function
    End If

    If r.Range.Tables.count = 0 Then
        Debug.Print "[IsValidRow] Row is not part of any table."
        IsValidRowForInsertion = False
        Exit Function
    End If

    If Len(Trim(r.Range.text)) = 0 Then
        Debug.Print "[IsValidRow] Row has no visible text."
        IsValidRowForInsertion = False
        Exit Function
    End If

    If expectedRegionTitle <> "" Then
        Dim cc As ContentControl
        For Each cc In r.Range.contentControls
            If ParentContentControlTitled(expectedRegionTitle, cc) Then
                IsValidRowForInsertion = True
                Exit Function
            End If
        Next cc
        Debug.Print "[IsValidRow] Row is not inside content control titled '" & expectedRegionTitle & "'."
        IsValidRowForInsertion = False
        Exit Function
    End If

    ' Fallback: all general checks passed
    IsValidRowForInsertion = True
    Exit Function

ErrorHandler:
    Debug.Print "[ERROR] IsValidRowForInsertion: " & Err.Description
    IsValidRowForInsertion = False
End Function

