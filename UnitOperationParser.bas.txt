Attribute VB_Name = "UnitOperationParser"
' === Module: UnitOperationParser ===
Option Explicit

Public Function ParseProcessDescription() As Collection
    Dim sourceTable As Table
    Dim parsedUnits As Collection
    Dim currentRow As row
    Dim unitOp As clsUnitOperation
    Dim rowIndex As Long

    Set parsedUnits = New Collection
    Set sourceTable = FindContentControlByTitle("ProcessDescription").Range.Tables(1)

    If sourceTable Is Nothing Then
        Debug.Print "[ERROR] 'ProcessDescription' table not found."
        Set ParseProcessDescription = parsedUnits
        Exit Function
    End If

    Debug.Print "[Parser] Starting ParseProcessDescription: " & sourceTable.rows.count & " rows"

    For rowIndex = 1 To sourceTable.rows.count
        Set currentRow = sourceTable.rows(rowIndex)

        Dim ccID As ContentControl
        Set ccID = FetchContentControlFromCell(currentRow.Cells(1))
        If ccID Is Nothing Then GoTo nextRow

        Set unitOp = New clsUnitOperation
        Set unitOp.sourceRow = currentRow
        unitOp.Step = Trim(ccID.Range.text)
        unitOp.tag = ccID.tag
        unitOp.title = ccID.title
        unitOp.id = Left(ccID.tag, 5)
        unitOp.InstructionText = ExtractInstructionText(currentRow.Cells(2))


        Dim cc As ContentControl
        For Each cc In currentRow.Cells(2).Range.contentControls
        ' Carefully check this list, incomplete!
            Select Case LCase(Trim(cc.title))
                Case "input": unitOp.AddInput Trim(cc.Range.text), cc.tag
                Case "waste": unitOp.AddOutput Trim(cc.Range.text), cc.tag
                Case "reactor", "upstreamreactor": unitOp.AddEquipment Trim(cc.Range.text), cc.tag
                Case Else
                    On Error GoTo ParamError
                    unitOp.AddParameter Trim(cc.title), Trim(cc.Range.text), cc.tag
                    On Error GoTo 0
            End Select
        Next cc

        parsedUnits.Add unitOp
nextRow:
    Next rowIndex

    Debug.Print "[Parser] Completed. Parsed " & parsedUnits.count & " unit operations."
    Set ParseProcessDescription = parsedUnits
    Exit Function

ParamError:
    Debug.Print "[ERROR] AddParameter failed in row " & rowIndex & ": " & Err.Description
    Err.Clear
    Resume Next
End Function


Private Function ExtractInstructionText(cell As cell) As String
    ExtractInstructionText = Trim(cell.Range.text)
End Function
