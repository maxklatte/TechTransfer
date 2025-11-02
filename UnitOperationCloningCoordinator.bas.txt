Attribute VB_Name = "UnitOperationCloningCoordinator"
' Module: UnitOperationCloningCoordinator
' Purpose: Provides preview functionality for cloning unit operations before executing ExecuteCloningWorkflow

Option Explicit

' === PUBLIC ENTRY POINT ===
Public Function GetClonePreview(allUnits As Collection, startIndex As Long, endIndex As Long, insertAfterIndex As Long) As Collection
    Dim preview As New Collection
    Dim i As Long
    Dim sourceIndexes() As Variant
    Dim unit As clsUnitOperation

    ' Build source indexes
    ReDim sourceIndexes(0 To endIndex - startIndex)
    For i = 0 To UBound(sourceIndexes)
        sourceIndexes(i) = startIndex + i
    Next i

    ' Define bounded preview window (2 before and 2 after insertion point)
    Dim lowerBound As Long, upperBound As Long
    lowerBound = SafeMax(1, insertAfterIndex - 2)
    upperBound = SafeMin(allUnits.count, insertAfterIndex + 2)

    ' Loop through preview window
    For i = lowerBound To upperBound
        Set unit = allUnits(i)
        preview.Add FormatPreviewLine(unit, i)

        ' Inject clone preview after the anchor
        If i = insertAfterIndex Then
            preview.Add "    --- CLONE INSERTION START ---"

            Dim j As Long
            Dim cloneIndex As Long
            For j = LBound(sourceIndexes) To UBound(sourceIndexes)
                cloneIndex = CLng(sourceIndexes(j))
                preview.Add "    [CLONE] " & FormatPreviewLine(allUnits(cloneIndex), cloneIndex)
            Next j

            preview.Add "    --- CLONE INSERTION END ---"
        End If
    Next i

    Set GetClonePreview = preview
End Function



' === HELPER FUNCTIONS ===
Private Function FormatPreviewLine(unit As clsUnitOperation, index As Long) As String
    On Error Resume Next
    Dim summary As String
    summary = unit.ToString()
    If summary = "" Then summary = "(Untitled unit operation)"
    FormatPreviewLine = index & ": " & summary
End Function

Private Function SafeMax(val1 As Long, val2 As Long) As Long
    If val1 > val2 Then
        SafeMax = val1
    Else
        SafeMax = val2
    End If
End Function

Private Function SafeMin(val1 As Long, val2 As Long) As Long
    If val1 < val2 Then
        SafeMin = val1
    Else
        SafeMin = val2
    End If
End Function

