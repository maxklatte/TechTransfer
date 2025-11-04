Attribute VB_Name = "UnitOperationHandler"
' Global collection to store generated IDs
Private generatedIDs As Collection

' === ID Registry Initialization ===
Public Sub InitializeIDRegistryFromIDs(idList As Variant)
    Set generatedIDs = New Collection
    Dim id As Variant
    On Error Resume Next
    For Each id In idList
        generatedIDs.Add id, id
    Next id
    On Error GoTo 0
End Sub

Public Sub Renumber()
 Application.ScreenUpdating = False
    On Error GoTo ErrorHandler
    
    Debug.Print "Starting Renumber..."
    
    ' Call custom procedures to renumber
    Call NumberIPCs
    Call NumberUnitOperations
    GetGlobalCompoundCollection.SortByStepNumber
    GetGlobalWasteCollection.SortByStepNumber
 Application.ScreenUpdating = True
    Exit Sub ' Ensure the function exits normally when no error occurs

ErrorHandler:
 Application.ScreenUpdating = True
    Debug.Print "Error in Renumber: " & Err.Description
    MsgBox "An error occurred while renumbering. Please check the debug output.", vbCritical, "Error"
    Err.Clear
    Resume Next ' Resume execution after the error
    
End Sub



Public Sub NumberUnitOperations()
    On Error GoTo ErrorHandler
    ' Debug.Print "Starting NumberUnitOperations..."

    Dim ccProcessDescription As ContentControl
    Dim targetTable As Table
    Dim targetRow As row
    Dim cell As cell
    Dim rowIndex As Integer

    ' Locate the content control titled "ProcessDescription"
    For Each ccProcessDescription In ThisDocument.contentControls
        If ccProcessDescription.title = "ProcessDescription" Then
            ' Debug.Print "Content control 'ProcessDescription' found."
            
            ' Check if the content control contains a table
            If ccProcessDescription.Range.Tables.count > 0 Then
                Set targetTable = ccProcessDescription.Range.Tables(1)
                ' Debug.Print "Table found inside 'ProcessDescription' content control."

                ' Iterate through each row in the table
                rowIndex = 1 ' Start numbering from 1
                For Each targetRow In targetTable.rows
                    Set cell = targetRow.Cells(1)
                    If cell.Range.contentControls.count > 0 Then
                        ' Update the first content control in the cell with the row index
                        cell.Range.contentControls(1).Range.text = CStr(rowIndex)
                        ' Debug.Print "Updated row " & targetRow.index & " with number: " & rowIndex
                        rowIndex = rowIndex + 1
                    Else
                        Debug.Print "No content control found in the first cell of row " & targetRow.index
                    End If
                Next targetRow

                ' Debug.Print "NumberUnitOperations completed successfully."

                ' Call SyncNumbersFromProcessDescription
                ' Debug.Print "Calling UtilityTableHandler.SyncNumbersFromProcessDescription..."
                On Error Resume Next
                ' Call UtilityTableHandler.SyncNumbersFromProcessDescription("Bill of Materials", "BOM")
                ' Call UtilityTableHandler.SyncNumbersFromProcessDescription("Waste Streams", "WS")
                ' Call UtilityTableHandler.SyncNumbersFromProcessDescription("IPCs", "IPC")
                ' Call UtilityTableHandler.SyncNumbersFromProcessDescription("IPCs", "PI")
                If Err.number <> 0 Then
                    Debug.Print "Error in SyncNumbersFromProcessDescription: " & Err.Description
                    Err.Clear
                End If
                On Error GoTo ErrorHandler

                Exit Sub
            Else
                Debug.Print "Error: No table found inside 'ProcessDescription' content control."
                MsgBox "The 'ProcessDescription' content control does not contain a table.", vbExclamation, "Error"
                Exit Sub
            End If
        End If
    Next ccProcessDescription

    ' If the content control titled "ProcessDescription" was not found
    Debug.Print "Error: Content control 'ProcessDescription' not found."
    MsgBox "Content control 'ProcessDescription' not found in the document.", vbExclamation, "Error"
    Exit Sub

ErrorHandler:
    Debug.Print "Error in NumberUnitOperations at row index " & rowIndex & ": " & Err.Description
    MsgBox "An error occurred while numbering unit operations. Please check the debug output.", vbCritical, "Error"
    Err.Clear
End Sub


' Bekannte Bugs: Wenn der Cursor am Ende in einem Drop Down ist, gibt es einen Fehler bei der Erstellung der BOM/etc. Liste:
'

Sub InitializeUnitOperation(ByVal targetRow As row)
    On Error GoTo ErrorHandler
    Debug.Print "Starting UnitOperationInitializer..."
    
    ' Disable screen updates for performance
    Application.ScreenUpdating = False
    
    Dim columnMap As Object
    Dim uniqueRowID As String
    Dim contentControlCounter As Integer
    Dim cell As cell
    Dim cc As ContentControl
    Dim previousRow As row
    Dim parentTable As Table

    ' Ensure global collection to track IDs is initialized
    If generatedIDs Is Nothing Then
        Set generatedIDs = New Collection
    End If

    ' Access the parent table
    Set parentTable = targetRow.Range.Tables(1)

    ' Ensure we can access the previous row
    If targetRow.index <= 1 Then
        Debug.Print "No previous row available for UnitOperationInitializer."
        Exit Sub
    End If

    ' Get the previous row
    Set previousRow = parentTable.rows(targetRow.index - 1)
    'Debug.Print "Processing previous row at index: " & previousRow.Index

    ' Collect all existing IDs in the table and add them to the collection
    CollectExistingIDs parentTable

    ' Generate a truly unique 5-digit ID for the row
    uniqueRowID = GenerateUniqueRowID()
    'Debug.Print "Generated unique row ID: " & uniqueRowID

    ' Initialize content control counter
    contentControlCounter = 1

    ' Assign unique tags to all content controls in the row
    For Each cell In previousRow.Cells
        If cell.Range.contentControls.count > 0 Then
            For Each cc In cell.Range.contentControls
                ' Assign a unique tag
                cc.tag = uniqueRowID & format(contentControlCounter, "00")
                'Debug.Print "Assigned tag: " & cc.Tag & " to content control in cell (Row: " & previousRow.Index & ", Column: " & cell.ColumnIndex & ")"
                contentControlCounter = contentControlCounter + 1
            Next cc
        Else
            Debug.Print "No content controls found in cell (Row: " & previousRow.index & ", Column: " & cell.columnIndex & ")"
        End If
    Next cell

    ' Renumber Unit Operations, so that the new numbers are available to be implemented in summary table "BOM" and "WS"
   '  Call UnitOperationHandler.NumberUnitOperations
    Call Renumber

    ' Process compound content controls
    'For Each cell In previousRow.Cells
    '    If cell.Range.contentControls.Count > 0 Then
    '        For Each cc In cell.Range.contentControls
    '            If LCase(cc.title) = "compound" Then
    '                Debug.Print "Compound content control found with tag: " & cc.tag
    '                Call UtilityTableHandler.DisplayDetails(cc, "Bill of Materials", "BOM")
    '                Call UtilityTableHandler.SyncNumbersFromProcessDescription("Bill of Materials", "BOM")
    '            End If
    '            If LCase(cc.title) = "stream" Then
    '                Call UtilityTableHandler.DisplayDetails(cc, "Waste Streams", "WS")
    '                Call UtilityTableHandler.SyncNumbersFromProcessDescription("Waste Streams", "WS")
    '            End If
    '            If LCase(cc.title) = "inprocesscontrol" Then
    '                Debug.Print "LCase(cc.Title):"; LCase(cc.title)
    '                Call NumberIPCs
    '                Call UtilityTableHandler.DisplayDetails(cc, "IPCs", "IPC")
    '                Call UtilityTableHandler.SyncNumbersFromProcessDescription("IPCs", "IPC")
    '            End If
    '            If LCase(cc.title) = "processindicator" Then
    '                Debug.Print "LCase(cc.Title):"; LCase(cc.title)
    '                Call NumberIPCs
    '                Call UtilityTableHandler.DisplayDetails(cc, "IPCs", "PI")
    '                Call UtilityTableHandler.SyncNumbersFromProcessDescription("IPCs", "PI")
    '            End If
    '        Next cc
    '    End If
    'Next cell

    'Debug.Print "UnitOperationInitializer completed successfully."

    ' Re-enable screen updates
    Application.ScreenUpdating = True

    Exit Sub

ErrorHandler:
    ' Re-enable screen updates in case of error
    Application.ScreenUpdating = True
    Debug.Print "Error in UnitOperationInitializer: " & Err.Description
    Err.Clear
End Sub


Private Sub CollectExistingIDs(ByVal parentTable As Table)
    Dim rowIndex As Integer
    Dim cell As cell
    Dim cc As ContentControl

    'Debug.Print "Collecting existing IDs from the table..."

    ' Iterate through all rows in the table
    For rowIndex = 2 To parentTable.rows.count ' Skip the header row
        Set cell = parentTable.rows(rowIndex).Cells(1) ' First column of each row
        If cell.Range.contentControls.count > 0 Then
            Set cc = cell.Range.contentControls(1) ' Assuming one content control per cell
            If Len(cc.tag) > 0 Then
                ' Add the existing ID to the collection
                On Error Resume Next
                generatedIDs.Add cc.tag, cc.tag ' Add as key to ensure uniqueness
                If Err.number = 0 Then
                    'Debug.Print "Existing ID added to collection: " & cc.tag
                Else
                    'Debug.Print "Duplicate ID found in collection: " & cc.tag
                    Err.Clear
                End If
                On Error GoTo 0
            Else
                Debug.Print "No tag found in content control at row " & rowIndex
            End If
        Else
            Debug.Print "No content control found in the first column of row " & rowIndex
        End If
    Next rowIndex
End Sub

Public Function GenerateUniqueRowID() As String
    On Error GoTo ErrorHandler
    Dim uniqueID As String
    Dim isUnique As Boolean

    ' Ensure randomization is initialized
    Randomize
    
    'To be replaced with clean class and singelton pattern
    If generatedIDs Is Nothing Then
    Set generatedIDs = New Collection
    End If

    Do
        ' Generate a random 5-digit ID
        uniqueID = format(Int((99999 - 10000 + 1) * Rnd + 10000), "00000")
        isUnique = True

        ' Check if the ID already exists in the collection
        On Error Resume Next
        generatedIDs.Add uniqueID, uniqueID ' Attempt to add the ID as a key
        If Err.number <> 0 Then
            isUnique = False ' Duplicate detected
            Err.Clear
        End If
        On Error GoTo 0
    Loop Until isUnique

    GenerateUniqueRowID = uniqueID
    Exit Function

ErrorHandler:
    Debug.Print "Error in GenerateUniqueRowID: " & Err.Description
    Err.Clear
    GenerateUniqueRowID = "00000" ' Fallback in case of error
End Function

Public Sub NumberIPCs()
    On Error GoTo ErrorHandler
    Debug.Print "Starting NumberIPCs..."

Dim ccProcessDescription As ContentControl
Dim cc As ContentControl
Dim ipcCounter As Integer
Dim piCounter As Integer

' Initialize counters for IPC and PI IDs
ipcCounter = 1
piCounter = 1

' Locate the content control titled "ProcessDescription"
For Each ccProcessDescription In ThisDocument.contentControls
    If ccProcessDescription.title = "ProcessDescription" Then
        'Debug.Print "Content control 'ProcessDescription' found."
        
        ' Iterate through all content controls within the ProcessDescription range
        For Each cc In ccProcessDescription.Range.contentControls
            Select Case LCase(cc.title)
                Case "ipc id"
                    ' For IPC IDs, use a numeric counter with prefix "IPC-"
                    cc.Range.text = "IPC-" & ipcCounter
                    'Debug.Print "Assigned IPC-" & ipcCounter & " to IPC ID"
                    'Call UtilityTableHandler.UpdateDetails(cc, "IPCs", "IPC")
                    ipcCounter = ipcCounter + 1
                    
                Case "pi id"
                    ' For PI IDs, use ConvertNumberToLetter for a letter-based counter with prefix "PI-"
                    cc.Range.text = "PI-" & ConvertNumberToLetter(piCounter)
                    'Debug.Print "Assigned PI-" & ConvertNumberToLetter(piCounter) & " to PI ID"
                    'Call UtilityTableHandler.UpdateDetails(cc, "IPCs", "PI")
                    piCounter = piCounter + 1
            End Select
        Next cc

        'Debug.Print "NumberIPCs completed successfully."
        Exit Sub
    End If
Next ccProcessDescription

Debug.Print "Error: Content control 'ProcessDescription' not found."
MsgBox "Content control 'ProcessDescription' not found in the document.", vbExclamation, "Error"
Exit Sub


ErrorHandler:
    Debug.Print "Error in NumberIPCs: " & Err.Description
    MsgBox "An error occurred while numbering IPCs. Please check the debug output.", vbCritical, "Error"
    Err.Clear
End Sub


Private Function ConvertNumberToLetter(ByVal num As Integer) As String
    ' Convert a number to a corresponding letter (1 -> A, 2 -> B, ..., 26 -> Z, 27 -> AA, etc.)
    Dim letters As String
    letters = ""

    Do While num > 0
        letters = Chr(((num - 1) Mod 26) + 65) & letters
        num = (num - 1) \ 26
    Loop

    ConvertNumberToLetter = letters
End Function

Public Sub PruneAllCollections()
    On Error GoTo ErrorHandler

    Dim compounds As CompoundCollection
    Dim wastes As WasteCollection
    Dim t0 As Double, t1 As Double, t2 As Double, t3 As Double

    t0 = Timer

    Set compounds = GetGlobalCompoundCollection()
    t1 = Timer
    compounds.PruneOrphaned
    t2 = Timer

    Set wastes = GetGlobalWasteCollection()
    wastes.PruneOrphaned
    t3 = Timer

    Debug.Print "TIMING (PruneAllCollections):"
    Debug.Print "  Get Compounds Time: " & format(t1 - t0, "0.000") & " sec"
    Debug.Print "  Compounds Prune Time: " & format(t2 - t1, "0.000") & " sec"
    Debug.Print "  Wastes Prune Time: " & format(t3 - t2, "0.000") & " sec"
    Debug.Print "  Total Time: " & format(t3 - t0, "0.000") & " sec"

    Exit Sub

ErrorHandler:
    Debug.Print "Error in PruneAllCollections: " & Err.Description
    MsgBox "Failed to prune collections: " & Err.Description, vbCritical
End Sub


