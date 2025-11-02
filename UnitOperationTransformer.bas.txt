Attribute VB_Name = "UnitOperationTransformer"
' Module: UnitOperationsHelpers
Option Explicit

' === Clone and Render Workflow ===

Public Sub ExecuteCloningWorkflow(startIndex As Long, endIndex As Long, insertAfterIndex As Long)
    Dim allUnits As Collection
    Dim clones As Collection
    Dim updatedUnits As Collection
    Dim sourceIndexes() As Variant
    Dim anchorID As String
    Dim i As Long

    On Error GoTo ErrorHandler
    Debug.Print "[Workflow] Begin: ExecuteCloningWorkflow(" & startIndex & ", " & endIndex & ", " & insertAfterIndex & ")"

    ' === Step 1: Parse current unit operations ===
    Set allUnits = ParseProcessDescription()
    If allUnits Is Nothing Or allUnits.count = 0 Then
        Err.Raise vbObjectError + 1001, , "[Workflow] No unit operations found in document."
    End If
    Debug.Print "[Workflow] Parsed " & allUnits.count & " unit operations."

    ' === Step 2: Validate input indexes ===
    If startIndex < 1 Or endIndex > allUnits.count Or insertAfterIndex > allUnits.count Then
        Err.Raise vbObjectError + 1002, , "[Workflow] Invalid index range."
    End If
    If endIndex < startIndex Then
        Err.Raise vbObjectError + 1003, , "[Workflow] endIndex cannot be less than startIndex."
    End If

    ' === Step 3: Prepare sourceIndexes array ===
    ReDim sourceIndexes(0 To endIndex - startIndex)
    For i = 0 To UBound(sourceIndexes)
        sourceIndexes(i) = startIndex + i
    Next i
    Debug.Print "[Workflow] Preparing to clone unit ops: " & Join(sourceIndexes, ", ")

    ' === Step 4: Clone the selected unit operations ===
    Set clones = CloneUnits(allUnits, sourceIndexes)
    If clones Is Nothing Or clones.count = 0 Then
        Err.Raise vbObjectError + 1004, , "[Workflow] CloneUnits returned no clones."
    End If
    Debug.Print "[Workflow] Cloned " & clones.count & " unit operations."

    ' === Step 5: Render them visually after anchor ===
    anchorID = allUnits(insertAfterIndex).id
    Call RenderClonedUnitOperations(clones, anchorID)
    Debug.Print "[Workflow] Rendered clones after ID: " & anchorID
    Debug.Print "[NOTE] Row highlight is applied for debug visibility — remove for production."

    ' === Step 6: Update the virtual data model ===
    Set updatedUnits = InsertUnits(allUnits, clones, insertAfterIndex)
    Debug.Print "[Workflow] Inserted clones after index " & insertAfterIndex & ". Final count: " & updatedUnits.count

    Debug.Print "[Workflow] End: ExecuteCloningWorkflow"
    Exit Sub

ErrorHandler:
    Debug.Print "[ERROR] ExecuteCloningWorkflow: " & Err.Description
    Err.Clear
    Resume Next
End Sub

Public Function CloneUnits(allUnits As Collection, sourceIndexes As Variant) As Collection
    Dim clonedUnits As Collection
    Set clonedUnits = New Collection

    Dim idx As Variant
    Dim clonedUnit As clsUnitOperation

    For Each idx In sourceIndexes
        If idx >= 1 And idx <= allUnits.count Then
            Set clonedUnit = allUnits(idx).Clone(GenerateUniqueRowID())
            clonedUnits.Add clonedUnit
        Else
            Debug.Print "[WARN] sourceIndex out of range: " & idx
        End If
    Next idx

    Set CloneUnits = clonedUnits
End Function

Public Function InsertUnits(originalUnits As Collection, pinsertUnits As Collection, targetIndex As Long) As Collection
    Dim result As Collection
    Set result = New Collection

    Dim i As Long

    ' Add units before the target index
    For i = 1 To targetIndex
        result.Add originalUnits(i)
    Next i

    ' Add the units to be inserted
    Dim unit As clsUnitOperation
    For Each unit In pinsertUnits
        result.Add unit
    Next unit

    ' Add the remaining original units
    For i = targetIndex + 1 To originalUnits.count
        result.Add originalUnits(i)
    Next i

    Set InsertUnits = result
End Function

Public Function CloneAndInsertUnits(allUnits As Collection, sourceIndexes As Variant, targetIndex As Long) As Collection
    Dim clonedUnits As Collection
    Set clonedUnits = CloneUnits(allUnits, sourceIndexes)

    Set CloneAndInsertUnits = InsertUnits(allUnits, clonedUnits, targetIndex)
End Function


