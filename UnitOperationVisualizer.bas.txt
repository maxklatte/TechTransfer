Attribute VB_Name = "UnitOperationVisualizer"
' === Module: UnitOperationVisualizer ===
Option Explicit

' === Entry Point ===
' === Module: RenderClonedOps_Rewrite ===


' === Description ===
' Renders a collection of cloned clsUnitOperation objects into the ProcessDescription table.
' Each cloned unit is inserted sequentially after the unit operation identified by anchorID.
' Tags are retagged immediately using each clone's .ID
'
' === Usage ===
' Call RenderClonedUnitOperations(clones, "49552")
'
Public Sub RenderClonedUnitOperations(clones As Collection, anchorID As String)
    Dim anchorCCs As Collection
    Dim anchorCC As ContentControl
    Dim anchorRow As row
    Dim u As clsUnitOperation
    Dim lastRow As row

    On Error GoTo ErrorHandler
    Debug.Print "[Render] Starting RenderClonedUnitOperations at anchor ID: " & anchorID

    ' === Step 1: Locate anchor row ===
    Set anchorCCs = FindControlsByTagInRegion(anchorID & "01", "ProcessDescription")
    If anchorCCs.count = 0 Then Err.Raise vbObjectError + 601, "RenderClonedUnitOperations", "Anchor tag not found: " & anchorID & "01"

    Set anchorCC = anchorCCs(1)
    Set anchorRow = anchorCC.Range.rows(1)
    Set lastRow = anchorRow
    Debug.Print "[Render] Anchor row index: " & anchorRow.index

    ' === Step 2: Insert each clone after the last inserted ===
    For Each u In clones
        Debug.Print "[Render] Inserting clone for ID: " & u.id
        Set lastRow = CloneAndRetagRow(u.SourceID, lastRow, u.id, "ProcessDescription")
        If Not lastRow Is Nothing Then
            Debug.Print "[Render] Inserted and retagged new row for: " & u.id & " at index: " & lastRow.index
        Else
            Debug.Print "[ERROR] Failed to insert row for unitOp ID: " & u.id
        End If
    Next u

    Debug.Print "[Render] Completed RenderClonedUnitOperations."
    Exit Sub

ErrorHandler:
    Debug.Print "[ERROR] RenderClonedUnitOperations: " & Err.Description
End Sub


' === Row Cloning Utility ===
Private Function CloneRowBelow(sourceRow As row, insertAfter As row) As row
    Dim newRow As row
    Dim tbl As Table

    On Error GoTo ErrorHandler
    Set tbl = insertAfter.Parent
    If tbl Is Nothing Then
        Debug.Print "[ERROR] CloneRowBelow: insertAfter has no parent table"
        Set CloneRowBelow = Nothing
        Exit Function
    End If

    If insertAfter.Next Is Nothing Then
        Set newRow = tbl.rows.Add()
        Debug.Print "[CloneRow] Appended new row at end of table."
    Else
        Set newRow = tbl.rows.Add(BeforeRow:=insertAfter.Next)
        Debug.Print "[CloneRow] Inserted row before index: " & insertAfter.Next.index
    End If

    sourceRow.Range.Copy
    newRow.Range.PasteAndFormat wdFormatOriginalFormatting

    Set CloneRowBelow = newRow
    Exit Function

ErrorHandler:
    Debug.Print "[ERROR] CloneRowBelow failed: " & Err.Description
    Set CloneRowBelow = Nothing
End Function

' === Tag Update Utility ===
Public Sub UpdateTagsInRow(r As row, newID As String)
    Dim cell As cell
    Dim cc As ContentControl
    Dim oldTag As String, suffix As String, newTag As String

    On Error GoTo ErrorHandler
    Debug.Print "[Retag] Updating tags in row index: " & r.index & " with newID: " & newID

    ' Loop over all cells (some tags might be outside of cell 2)
    For Each cell In r.Cells
        For Each cc In cell.Range.contentControls
            oldTag = cc.tag
            If Len(oldTag) >= 2 Then
                suffix = Right(oldTag, 2)
                newTag = newID & suffix
                Debug.Print "[Retag] old: " & oldTag & " ? new: " & newTag
                cc.tag = newTag
            Else
                Debug.Print "[WARN] Skipped control with short/empty tag: " & oldTag
            End If
        Next cc
    Next cell

    Exit Sub

ErrorHandler:
    Debug.Print "[ERROR] UpdateTagsInRow: " & Err.Description
End Sub
Public Function CloneAndRetagRow(SourceID As String, insertAfter As row, newID As String, regionTitle As String) As row
    Dim ccList As Collection
    Dim cc As ContentControl
    Dim sourceRow As row
    Dim newRow As row

    On Error GoTo ErrorHandler
    Debug.Print "[CloneAndRetag] Start for sourceID: " & SourceID & ", newID: " & newID

    ' === Locate source tag (assumes XXXXX01 format) ===
    Set ccList = FindControlsByTagInRegion(SourceID & "01", regionTitle)
    If ccList.count = 0 Then Err.Raise vbObjectError + 510, "CloneAndRetagRow", "No content control with tag " & SourceID & "01 found in region '" & regionTitle & "'"

    If ccList.count > 1 Then
        Debug.Print "[WARN] Multiple controls found for " & SourceID & "01, using the first (LIMITATION)"
    End If

    Set cc = ccList(1)
    Set sourceRow = cc.Range.rows(1)

    If Not IsValidRowForInsertion(sourceRow, regionTitle) Then
        Err.Raise vbObjectError + 511, "CloneAndRetagRow", "Source row is invalid or not in region '" & regionTitle & "'"
    End If

    ' === Clone and retag ===
    Set newRow = InsertClonedRowAfter(sourceRow, insertAfter)
    If newRow Is Nothing Then
        Err.Raise vbObjectError + 512, "CloneAndRetagRow", "Row was not successfully inserted."
    End If

    UpdateTagsInRow newRow, newID
    'ApplyVisualHighlight newRow
    Set CloneAndRetagRow = newRow

    Debug.Print "[CloneAndRetag] Done for newID: " & newID & " ? row index: " & newRow.index
    Exit Function

ErrorHandler:
    Debug.Print "[ERROR] CloneAndRetagRow: " & Err.Description
    Set CloneAndRetagRow = Nothing
End Function


' === Visual Debugging Highlight ===
Private Sub ApplyVisualHighlight(targetRow As row)
    On Error Resume Next
    Dim c As cell
    For Each c In targetRow.Cells
        c.Shading.BackgroundPatternColor = wdColorLightOrange
    Next c
    Debug.Print "[Highlight] Row index " & targetRow.index & " shaded orange."
End Sub

' === Module: Test_CloneUnitOp2Once ===

' === Module: Test_CloneUnitOp2Once ===

Public Sub OLDTest_CloneUnitOp2Once()
    Dim ccList As Collection
    Dim ccAnchor As ContentControl
    Dim sourceRow As row
    Dim newRow As row
    Dim nextRow As row

    On Error GoTo ErrorHandler
    Debug.Print "[Test] Begin: Cloning Unit Operation 2 (ID: 49552)"

    ' === Step 1: Locate anchor control ===
    Set ccList = FindControlsByTagInRegion("4955201", "ProcessDescription")
    If ccList.count = 0 Then
        Debug.Print "[Test] ERROR: No control found with tag 4955201 in region ProcessDescription."
        Exit Sub
    End If

    Set ccAnchor = ccList(1)
    Set sourceRow = ccAnchor.Range.rows(1)
    Debug.Print "[Test] Found source row at index: " & sourceRow.index
    Debug.Print "[Test] Source row content: " & vbCrLf & sourceRow.Range.text

    ' === Step 2: Validate source row ===
    If Not IsValidRowForInsertion(sourceRow, "ProcessDescription") Then
        Debug.Print "[Test] ERROR: Source row is not valid for insertion."
        Exit Sub
    End If

    ' === Step 3: Clone below ===
    Set newRow = InsertClonedRowAfter(sourceRow, sourceRow)

    If newRow Is Nothing Then
        Debug.Print "[Test] ERROR: InsertClonedRowAfter returned Nothing."
    Else
        Debug.Print "[Test] SUCCESS: Cloned row inserted at index: " & newRow.index
        Debug.Print "[Test] Cloned row content (before highlight): " & vbCrLf & newRow.Range.text
        'ApplyVisualHighlight newRow

        ' === Verify pasted result ===
        Debug.Print "[Verify] newRow.Range.Text:" & vbCrLf & newRow.Range.text
        If Not sourceRow.Next Is Nothing Then
            Debug.Print "[Verify] sourceRow.Next.Range.Text:" & vbCrLf & sourceRow.Next.Range.text
        End If
        If Not newRow.Next Is Nothing Then
            Set nextRow = newRow.Next
            Debug.Print "[Test] Next row after cloned row (index " & nextRow.index & ") content:" & vbCrLf & nextRow.Range.text
        Else
            Debug.Print "[Test] No row follows cloned row."
        End If
    End If

    Debug.Print "[Test] End: Test_CloneUnitOp2Once complete."
    Exit Sub

ErrorHandler:
    Debug.Print "[ERROR] Test_CloneUnitOp2Once: " & Err.Description
End Sub


' === New, clean implementation ===
' InsertClonedRowAfter(sourceRow, insertAfter)
' Description:
'   Copies a source Word table row and inserts it after the specified row.
'   Uses InsertParagraphAfter + PasteAndFormat to let Word manage layout.
'   Returns the newly inserted row (insertAfter.Next).
'
' Best Practices:
' - Use only with validated source/target rows
' - Always call UpdateTagsInRow(newRow, newID) immediately after
' - Apply visual highlight for dev visibility
Public Function InsertClonedRowAfter(sourceRow As row, insertAfter As row) As row
    On Error GoTo ErrorHandler
    Debug.Print "[InsertClone] Copying row " & sourceRow.index & " to insert after row " & insertAfter.index

    sourceRow.Range.Copy
    insertAfter.Range.InsertParagraphAfter

    If insertAfter.Next Is Nothing Then
        Debug.Print "[InsertClone] WARNING: insertAfter.Next is Nothing — trying fallback via Table.Rows.Add"

        Dim tbl As Table
        If insertAfter.Range.Tables.count = 0 Then
            Debug.Print "[InsertClone] ERROR: No table found in insertAfter.Range"
            Set InsertClonedRowAfter = Nothing
            Exit Function
        End If

        Set tbl = insertAfter.Range.Tables(1)
        Dim addedRow As row
        Set addedRow = tbl.rows.Add
        addedRow.Range.PasteAndFormat wdFormatOriginalFormatting
        Set InsertClonedRowAfter = addedRow
        Debug.Print "[InsertClone] Fallback: New row added via tbl.Rows.Add at index: " & addedRow.index
        Exit Function
    End If

    insertAfter.Next.Range.PasteAndFormat wdFormatOriginalFormatting

    If Not insertAfter.Next Is Nothing Then
        ' Note to developer: This edge case needs additional work and may hint at a complete reinvention of the entire function, to be done later.
        Debug.Print "[InsertClone] New row appears at index: " & insertAfter.Next.index
        Set InsertClonedRowAfter = insertAfter.Next
    Else
        Debug.Print "[InsertClone] WARNING: Paste succeeded but .Next row is still Nothing."
        Set InsertClonedRowAfter = Nothing
    End If
    Exit Function

ErrorHandler:
    Debug.Print "[ERROR] InsertClonedRowAfter failed: " & Err.Description
    Set InsertClonedRowAfter = Nothing
End Function


Public Sub Test_CloneUnitOp2Once()
    Dim ccList As Collection
    Dim ccAnchor As ContentControl
    Dim sourceRow As row
    Dim newRow As row
    Dim nextRow As row

    On Error GoTo ErrorHandler
    Debug.Print "[Test] Begin: Cloning Unit Operation 2 (ID: 49552)"

    ' === Step 1: Locate anchor control ===
    Set ccList = FindControlsByTagInRegion("4955201", "ProcessDescription")
    If ccList.count = 0 Then
        Debug.Print "[Test] ERROR: No control found with tag 4955201 in region ProcessDescription."
        Exit Sub
    End If

    Set ccAnchor = ccList(1)
    Set sourceRow = ccAnchor.Range.rows(1)
    Debug.Print "[Test] Found source row at index: " & sourceRow.index
    Debug.Print "[Test] Source row content: " & vbCrLf & sourceRow.Range.text

    ' === Step 2: Validate source row ===
    If Not IsValidRowForInsertion(sourceRow, "ProcessDescription") Then
        Debug.Print "[Test] ERROR: Source row is not valid for insertion."
        Exit Sub
    End If

    ' === Step 3: Clone below ===
    Set newRow = InsertClonedRowAfter(sourceRow, sourceRow)

    If newRow Is Nothing Then
        Debug.Print "[Test] ERROR: InsertClonedRowAfter returned Nothing."
    Else
        Debug.Print "[Test] SUCCESS: Cloned row inserted at index: " & newRow.index
        Debug.Print "[Test] Cloned row content (before highlight): " & vbCrLf & newRow.Range.text
        'ApplyVisualHighlight newRow

        ' === Step 4: Retag with dummy ID ===
        Call UpdateTagsInRow(newRow, "00000")

        ' === Verify pasted result ===
        Debug.Print "[Verify] newRow.Range.Text:" & vbCrLf & newRow.Range.text
        If Not sourceRow.Next Is Nothing Then
            Debug.Print "[Verify] sourceRow.Next.Range.Text:" & vbCrLf & sourceRow.Next.Range.text
        End If
        If Not newRow.Next Is Nothing Then
            Set nextRow = newRow.Next
            Debug.Print "[Test] Next row after cloned row (index " & nextRow.index & ") content:" & vbCrLf & nextRow.Range.text
        Else
            Debug.Print "[Test] No row follows cloned row."
        End If
    End If

    Debug.Print "[Test] End: Test_CloneUnitOp2Once complete."
    Exit Sub

ErrorHandler:
    Debug.Print "[ERROR] Test_CloneUnitOp2Once: " & Err.Description
End Sub

Public Sub Test_CloneAndRetagRow_Once()
    Dim anchorCCs As Collection
    Dim anchorCC As ContentControl
    Dim anchorRow As row
    Dim clonedRow As row
    Dim p As Paragraph

    On Error GoTo ErrorHandler
    Debug.Print "[Test] Begin: Test_CloneAndRetagRow_Once"

    ' === Find anchor row (use UnitOp ID 49552 as example) ===
    Set anchorCCs = FindControlsByTagInRegion("4955201", "ProcessDescription")
    If anchorCCs.count = 0 Then
        Debug.Print "[Test] ERROR: Anchor content control not found."
        Exit Sub
    End If

    Set anchorCC = anchorCCs(1)
    Set anchorRow = anchorCC.Range.rows(1)
    Debug.Print "[Test] Anchor row index: " & anchorRow.index

    ' === Inspect source row paragraphs ===
    Debug.Print "[Inspect] Paragraphs in source row:"
    For Each p In anchorRow.Range.Paragraphs
        Debug.Print "[Para] Len=" & Len(p.Range.text) & " ? '" & Replace(p.Range.text, vbCr, "¶") & "'"
    Next p

    ' === Call test ===
    Set clonedRow = CloneAndRetagRow("49552", anchorRow, "00001", "ProcessDescription")

    If clonedRow Is Nothing Then
        Debug.Print "[Test] ERROR: CloneAndRetagRow returned Nothing."
    Else
        Debug.Print "[Test] SUCCESS: Cloned + retagged row at index: " & clonedRow.index
    End If

    Debug.Print "[Test] End: Test_CloneAndRetagRow_Once"
    Exit Sub

ErrorHandler:
    Debug.Print "[ERROR] Test_CloneAndRetagRow_Once: " & Err.Description
End Sub

' === Module: Test_CloneMultipleOps ===
Public Sub Test_CloneMultipleOps_Manual()
    Dim anchorCCs As Collection
    Dim anchorCC As ContentControl
    Dim anchorRow As row
    Dim lastRow As row

    On Error GoTo ErrorHandler
    Debug.Print "[Test] Begin: Test_CloneMultipleOps_Manual"

    ' === Step 1: Locate anchor ===
    Set anchorCCs = FindControlsByTagInRegion("4955201", "ProcessDescription") ' UnitOp 2
    If anchorCCs.count = 0 Then
        Debug.Print "[Test] ERROR: Anchor content control not found."
        Exit Sub
    End If
    Set anchorCC = anchorCCs(1)
    Set anchorRow = anchorCC.Range.rows(1)
    Debug.Print "[Test] Anchor row index: " & anchorRow.index

    ' === Step 2: Chain clone and retag ===
    Set lastRow = CloneAndRetagRow("49552", anchorRow, "10001", "ProcessDescription") ' UnitOp 2 ? copy of itself
    Set lastRow = CloneAndRetagRow("22250", lastRow, "10002", "ProcessDescription") ' UnitOp 3
    Set lastRow = CloneAndRetagRow("63834", lastRow, "10003", "ProcessDescription") ' UnitOp 4
    Set lastRow = CloneAndRetagRow("30637", lastRow, "10004", "ProcessDescription") ' UnitOp 5

    Debug.Print "[Test] End: Test_CloneMultipleOps_Manual"
    Exit Sub

ErrorHandler:
    Debug.Print "[ERROR] Test_CloneMultipleOps_Manual: " & Err.Description
End Sub
' === Module: Test_CloneMultipleOps ===

Public Sub Test_CloneEdgeCases()
    Dim row1 As row, row5 As row
    Dim cloneA As row, cloneB As row

    On Error GoTo ErrorHandler
    Debug.Print "[Test] Begin: Test_CloneEdgeCases"

    ' === Find row for UnitOp 1 ===
    Set row1 = FindRowByTag("4979001", "ProcessDescription")
    If row1 Is Nothing Then Err.Raise vbObjectError + 601, "Test", "Anchor row for UnitOp 1 not found"
    Debug.Print "[Test] Found UnitOp 1 at row: " & row1.index

    ' === Insert UnitOp 5 after UnitOp 1 ===
    Set cloneA = CloneAndRetagRow("30637", row1, "20001", "ProcessDescription")
    Debug.Print "[Test] Inserted UnitOp 5 after UnitOp 1 ? row index: " & cloneA.index

    ' === Refresh reference to UnitOp 5's row ===
    Set row5 = FindRowByTag("3063701", "ProcessDescription")
    If row5 Is Nothing Then Err.Raise vbObjectError + 602, "Test", "Anchor row for UnitOp 5 not found after shift"
    Debug.Print "[Test] Found UnitOp 5 at row: " & row5.index

    ' === Insert UnitOp 1 after UnitOp 5 ===
    Set cloneB = CloneAndRetagRow("49790", row5, "20002", "ProcessDescription")
    Debug.Print "[Test] Inserted UnitOp 1 after UnitOp 5 ? row index: " & cloneB.index

    Debug.Print "[Test] End: Test_CloneEdgeCases"
    Exit Sub

ErrorHandler:
    Debug.Print "[ERROR] Test_CloneEdgeCases: " & Err.Description
End Sub


Public Function FindRowByTag(tagId As String, region As String) As row
    Dim ccList As Collection, cc As ContentControl
    Set ccList = FindControlsByTagInRegion(tagId, region)
    If ccList.count > 0 Then
        Set cc = ccList(1)
        Set FindRowByTag = cc.Range.rows(1)
    Else
        Set FindRowByTag = Nothing
    End If
End Function

' === Module: Test_RenderClones ===

Public Sub Test_RenderClonedUnitOps_After2()
    Dim allUnits As Collection
    Dim clones As Collection
    Dim sourceIndexes(2) As Variant

    On Error GoTo ErrorHandler
    Debug.Print "[Test] Begin: Render cloned units after UnitOp 2"

    ' === Step 1: Parse existing operations ===
    Set allUnits = ParseProcessDescription()

    If allUnits.count < 3 Then
        Debug.Print "[Test] ERROR: Need at least 3 unit operations to test."
        Exit Sub
    End If

    ' === Step 2: Pick UnitOps 3, 4, 5 to clone ===
    sourceIndexes(0) = 3
    sourceIndexes(1) = 4
    sourceIndexes(2) = 5

    ' === Step 3: Clone into memory ===
    ' NEEDS WORK, this does not return the clones, but whole collection including clones. We are now at the point, where we need to correctly get the visual render and the virtual copying process aligned. For this we need to discuss and seperate responsibilties. What we need is
    ' 1. a parsed collection as start point, insertion point, clones
    ' 2. visualization of the clones (recieving insertion point)
    ' 3. optional return the fully cloned colection digital twin for checking, it does not have a real point as state is preserved in wastecollection compound collection, the tags and our word table "process description"
    
    Set clones = CloneAndInsertUnits(allUnits, sourceIndexes, 2)

    ' === Step 4: Render visually into Word ===
    Call RenderClonedUnitOperations(clones, "49552") ' Insert after UnitOp 2

    Debug.Print "[Test] End: Test_RenderClonedUnitOps_After2"
    Exit Sub

ErrorHandler:
    Debug.Print "[ERROR] Test_RenderClonedUnitOps_After2: " & Err.Description
End Sub


