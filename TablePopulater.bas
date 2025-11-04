Attribute VB_Name = "TablePopulater"
' MASSIVE POTENTIAL FOR REFACTOR AND SPEED INCREASE HERE


'=== MODULE: TablePopulator ===
Option Explicit

' Populates the Bill of Materials (BOM) table from a CompoundCollection
Public Sub PopulateCompoundsTable(ByVal compounds As CompoundCollection)
    On Error GoTo ErrorHandler

    Dim startTime As Double
    Dim pruneTime As Double
    Dim clearTime As Double
    Dim loopStartTime As Double
    Dim loopEndTime As Double
    Dim totalTime As Double

    startTime = Timer ' Start total timing

    Dim ccBOM As ContentControl
    Dim bomTable As Table
    Dim row As row
    Dim cmp As Compound
    Dim i As Long

    Application.ScreenUpdating = False

    ' Locate BOM table content control
    For Each ccBOM In ThisDocument.contentControls
        If ccBOM.title = "Bill of Materials" Then
            If ccBOM.Range.Tables.count > 0 Then
                Set bomTable = ccBOM.Range.Tables(1)
                Exit For
            End If
        End If
    Next ccBOM

    If bomTable Is Nothing Then
        MsgBox "Error: Could not find BOM table.", vbExclamation
        Exit Sub
    End If

    ' Clear all rows except the header row
    ' === Refactored: Clear all rows except the header row ===
    While bomTable.rows.count > 1
        bomTable.rows(2).Delete
    Wend
    clearTime = Timer

    ' PruneOrphaned Compounds
    'compounds.PruneOrphaned
    pruneTime = Timer

    Dim ref As Compound
    Set ref = compounds.GetReferenceCompound()

    ' Loop through compounds and add to table
For i = 1 To compounds.count
    loopStartTime = Timer

    Set cmp = compounds.item(i)
    Set row = bomTable.rows.Add
    row.Range.Font.Bold = False

    Call InsertCompoundRow(row, cmp, ref)

    loopEndTime = Timer
Next i

    totalTime = Timer

    Debug.Print "=== PopulateCompoundsTable Timing ==="
    Debug.Print "  Table clear time:      " & format(clearTime - startTime, "0.000") & " sec"
    Debug.Print "  Prune time:            " & format(pruneTime - clearTime, "0.000") & " sec"
    Debug.Print "  Population time:       " & format(totalTime - pruneTime, "0.000") & " sec"
    Debug.Print "  TOTAL time:            " & format(totalTime - startTime, "0.000") & " sec"

    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Debug.Print "Error in PopulateCompoundsTable: " & Err.Description
    MsgBox "Failed to populate BOM table: " & Err.Description, vbCritical
    Application.ScreenUpdating = True
End Sub

Private Sub InsertCompoundRow(ByVal row As row, ByVal cmp As Compound, ByVal ref As Compound)
Dim equivText As String

If Not ref Is Nothing Then
    If Not ref.amount Is Nothing Then
        If ref.amount.MolarAmount <> 0 Then
            'Debug.Print "[DBG] Calculating equivText from ref + cmp molar amounts"
            equivText = SigStr(cmp.amount.GetCorrectedMolarAmount(cmp.Stoffdaten) / ref.amount.GetCorrectedMolarAmount(ref.Stoffdaten))
        Else
            Debug.Print "[DBG] cmp.amount.GetCorrectedMolarAmount(cmp.Stoffdaten) = 0"
            equivText = "---"
        End If
    Else
        Debug.Print "[DBG] ref.amount is Nothing"
        equivText = "---"
    End If
Else
    Debug.Print "[DBG] ref is Nothing"
    equivText = "---"
End If


    Dim typ As String
    typ = LCase(Trim(cmp.compoundType))
    If typ = "" Then typ = "solvent" ' Default to solvent

    Call InsertCellContent(row.Cells(1), CStr(GetStepNumberFromId(cmp.id)), "step_number", cmp.id & "BOM")
    Call InsertCellContent(row.Cells(2), cmp.Stoffdaten.productCode, "product_code", cmp.id & "BOM")
    Call InsertCellContent(row.Cells(3), cmp.Stoffdaten.title, "product_name", cmp.id & "BOM")

    
    If typ = "product" Then
    Call InsertCellContent(row.Cells(1), CStr(GetStepNumberFromId(cmp.id)), "step_number", cmp.id & "BOM")
    Call InsertCellContent(row.Cells(2), cmp.Stoffdaten.productCode, "product_code", cmp.id & "BOM")
    Call InsertCellContent(row.Cells(3), "Product", "product_name", cmp.id & "BOM")
    Call InsertCellContent(row.Cells(4), format(cmp.Stoffdaten.molarMass, "0.00"), "molecular_weight", cmp.id & "BOM")
    Call InsertCellContent(row.Cells(5), SigStr(cmp.amount.mass), "mass", cmp.id & "BOM")
    Call InsertCellContent(row.Cells(6), SigStr(cmp.amount.Volume), "volume", cmp.id & "BOM")
    Call InsertCellContent(row.Cells(7), SigStr(cmp.amount.GetCorrectedMolarAmount(cmp.Stoffdaten) * 1000), "molar_amount", cmp.id & "BOM")
    Call InsertCellContent(row.Cells(8), equivText, "molar_equivalents", cmp.id & "BOM")
    ElseIf typ = "reactant" Then
        Call InsertCellContent(row.Cells(4), format(cmp.Stoffdaten.molarMass, "0.00"), "molecular_weight", cmp.id & "BOM")
        Call InsertCellContent(row.Cells(5), SigStr(cmp.amount.mass), "mass", cmp.id & "BOM")
        Call InsertCellContent(row.Cells(6), SigStr(cmp.amount.Volume), "volume", cmp.id & "BOM")
        Call InsertCellContent(row.Cells(7), SigStr(cmp.amount.GetCorrectedMolarAmount(cmp.Stoffdaten) * 1000), "molar_amount", cmp.id & "BOM")
        Call InsertCellContent(row.Cells(8), equivText, "molar_equivalents", cmp.id & "BOM")
    ElseIf typ = "solvent" Then
        Call InsertCellContent(row.Cells(4), "---", "molecular_weight", cmp.id & "BOM")
        Call InsertCellContent(row.Cells(5), SigStr(cmp.amount.mass), "mass", cmp.id & "BOM")
        Call InsertCellContent(row.Cells(6), SigStr(cmp.amount.Volume), "volume", cmp.id & "BOM")
        Call InsertCellContent(row.Cells(7), "---", "molar_amount", cmp.id & "BOM")
        Call InsertCellContent(row.Cells(8), "---", "molar_equivalents", cmp.id & "BOM")
    ElseIf typ = "reagent" Then
        Call InsertCellContent(row.Cells(4), "---", "molecular_weight", cmp.id & "BOM")
        Call InsertCellContent(row.Cells(5), SigStr(cmp.amount.mass), "mass", cmp.id & "BOM")
        Call InsertCellContent(row.Cells(6), "---", "volume", cmp.id & "BOM")
        Call InsertCellContent(row.Cells(7), "---", "molar_amount", cmp.id & "BOM")
        Call InsertCellContent(row.Cells(8), "---", "molar_equivalents", cmp.id & "BOM")
    End If
End Sub




' Populates the Waste Streams table from a WasteCollection
Public Sub PopulateWasteTable(ByVal wastes As WasteCollection)
    On Error GoTo ErrorHandler

    Dim ccWS As ContentControl
    Dim wsTable As Table
    Dim row As row
    Dim w As Waste
    Dim i As Long

    Application.ScreenUpdating = False

    ' Locate Waste Streams table content control
    For Each ccWS In ThisDocument.contentControls
        If ccWS.title = "Waste Streams" Then
            If ccWS.Range.Tables.count > 0 Then
                Set wsTable = ccWS.Range.Tables(1)
                Exit For
            End If
        End If
    Next ccWS

    If wsTable Is Nothing Then
        MsgBox "Error: Could not find Waste Streams table.", vbExclamation
        Exit Sub
    End If

    ' === Refactored: Clear all rows except the header row ===
    While wsTable.rows.count > 1
        wsTable.rows(2).Delete
    Wend
        ' PruneOrphaned wastes before compiling the table
    'wastes.PruneOrphaned

    ' Loop through wastes and add to table
    For i = 1 To wastes.count
        Set w = wastes.item(i)

        ' Add new row at the end
        Set row = wsTable.rows.Add
        row.Range.Font.Bold = False

        ' Insert waste details into table
        Call InsertCellContent(row.Cells(1), CStr(GetStepNumberFromId(w.id)), "step_number", w.id & "WS")
        Call InsertCellContent(row.Cells(2), w.WasteType, "type", w.id & "WS")
        Call InsertCellContent(row.Cells(3), SigStr(w.mass), "weight", w.id & "WS")
        Call InsertCellContent(row.Cells(4), SigStr(w.Volume), "volume", w.id & "WS")
        Call InsertCellContent(row.Cells(5), IIf(Len(Trim(w.Description)) > 0, w.Description, "---"), "description", w.id & "WS")
    Next i

    'Debug.Print "Waste table populated with " & wastes.Count & " items."
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Debug.Print "Error in PopulateWasteTable: " & Err.Description
    MsgBox "Failed to populate Waste Streams table: " & Err.Description, vbCritical
    Application.ScreenUpdating = True
End Sub


'=== MODULE: MassBalanceTablePopulator ===
Public Sub PopulateMassBalanceTable()
    On Error GoTo ErrorHandler

    Dim t0 As Double, t1 As Double, t2 As Double, t3 As Double, tFinal As Double
    t0 = Timer ' Start total

    Dim ccMB As ContentControl
    Dim mbTable As Table
    Dim row As row
    Dim allEntries As Collection
    Dim compounds As CompoundCollection
    Dim wastes As WasteCollection
    Dim i As Long

    Set compounds = GetGlobalCompoundCollection()
    Set wastes = GetGlobalWasteCollection()

    Application.ScreenUpdating = False

    ' Locate Mass Balance table
    For Each ccMB In ThisDocument.contentControls
        If ccMB.title = "Mass Balance" Then
            If ccMB.Range.Tables.count > 0 Then
                Set mbTable = ccMB.Range.Tables(1)
                Exit For
            End If
        End If
    Next ccMB

    If mbTable Is Nothing Then
        MsgBox "Error: Could not find Mass Balance table.", vbExclamation
        Exit Sub
    End If

    ' === Refactored: Clear all rows except the header row ===
    While mbTable.rows.count > 1
        mbTable.rows(2).Delete
    Wend

    ' Prune orphaned data
    'compounds.PruneOrphaned
    'wastes.PruneOrphaned
    t1 = Timer

    ' Merge and tag entries
    Set allEntries = New Collection

    For i = 1 To compounds.count
        Dim cmp As Compound
        Set cmp = compounds.item(i)

        Dim cEntry As MassBalanceEntry
        Set cEntry = New MassBalanceEntry
        cEntry.Step = GetStepNumberFromId(cmp.id)
    If LCase(cmp.compoundType) = "product" Then
        cEntry.EntryType = "output"
    Else
        cEntry.EntryType = "compound"
    End If

        cEntry.title = cmp.Stoffdaten.title
        cEntry.UnitOperation = GetUnitOperationFromId(cmp.id)
        cEntry.Weight = cmp.amount.mass
        cEntry.Volume = cmp.amount.Volume

        allEntries.Add cEntry
    Next i

    For i = 1 To wastes.count
        Dim w As Waste
        Set w = wastes.item(i)

        Dim wEntry As MassBalanceEntry
        Set wEntry = New MassBalanceEntry
        wEntry.Step = GetStepNumberFromId(w.id)
        wEntry.EntryType = "waste"
        wEntry.title = IIf(Len(Trim(w.Description)) > 0, w.Description, "---")
        wEntry.UnitOperation = GetUnitOperationFromId(w.id)
        wEntry.Weight = w.mass
        wEntry.Volume = w.Volume

        allEntries.Add wEntry
    Next i
    t2 = Timer

    ' Sort entries by Step then EntryType (compound before waste)
    Call SortMassBalanceEntries(allEntries)
    t3 = Timer

    ' Populate table
    Dim totalWeight As Double: totalWeight = 0
    Dim totalVolume As Double: totalVolume = 0
    Dim entry As MassBalanceEntry

    For Each entry In allEntries
        Set row = mbTable.rows.Add
        row.Range.Font.Bold = False

        Dim deltaWeight As Double
        Dim deltaVolume As Double

        If entry.EntryType = "compound" Then
            deltaWeight = entry.Weight
            deltaVolume = entry.Volume
        Else
            deltaWeight = -entry.Weight
            deltaVolume = -entry.Volume
        End If

        totalWeight = totalWeight + deltaWeight
        totalVolume = totalVolume + deltaVolume

        'Debug.Print "Cell(1) content controls: " & row.Cells(1).Range.contentControls.count
        
        Call InsertCellContent(row.Cells(1), CStr(entry.Step), "step", entry.id)
        Call InsertCellContent(row.Cells(2), entry.UnitOperation, "unit_operation", entry.id)
        Call InsertCellContent(row.Cells(3), entry.title, "title", entry.id)
        Call InsertCellContent(row.Cells(4), SigStr(deltaWeight), "added_weight", entry.id)
        Call InsertCellContent(row.Cells(5), SigStr(deltaVolume), "added_volume", entry.id)
        Call InsertCellContent(row.Cells(6), SigStr(totalWeight), "total_weight", entry.id)
        Call InsertCellContent(row.Cells(7), SigStr(totalVolume), "total_volume", entry.id)
    Next entry

    tFinal = Timer

    Debug.Print "TIMING (Mass Balance Table):"
    Debug.Print "  Prune Time: " & format(t1 - t0, "0.000") & " sec"
    Debug.Print "  Entry Creation Time: " & format(t2 - t1, "0.000") & " sec"
    Debug.Print "  Sorting Time: " & format(t3 - t2, "0.000") & " sec"
    Debug.Print "  Table Population Time: " & format(tFinal - t3, "0.000") & " sec"
    Debug.Print "  Total Time: " & format(tFinal - t0, "0.000") & " sec"

    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Debug.Print "Error in PopulateMassBalanceTable: " & Err.Description
    MsgBox "Failed to populate Mass Balance table: " & Err.Description, vbCritical
    Application.ScreenUpdating = True
End Sub



'=== Helper: Sort by Step and Compound First ===
'=== Safe Sorting: Use Temporary Array to Sort MassBalanceEntries ===
Sub SortMassBalanceEntries(entries As Collection)
    On Error GoTo SortError
    
        If entries Is Nothing Or entries.count = 0 Then
        Debug.Print "[DBG] Sort skipped: entries collection is empty"
        Exit Sub
    End If

    Dim entryArray() As MassBalanceEntry
    Dim sortedCollection As New Collection
    Dim i As Long, j As Long

    ' Copy to array
    ReDim entryArray(1 To entries.count)
    For i = 1 To entries.count
        Set entryArray(i) = entries(i)
    Next i

    ' Bubble sort array safely
    For i = 1 To UBound(entryArray) - 1
        For j = i + 1 To UBound(entryArray)
            If entryArray(i).Step > entryArray(j).Step Or _
               (entryArray(i).Step = entryArray(j).Step And entryArray(i).EntryType = "waste" And entryArray(j).EntryType = "compound") Then

                Dim tmp As MassBalanceEntry
                Set tmp = entryArray(i)
                Set entryArray(i) = entryArray(j)
                Set entryArray(j) = tmp
            End If
        Next j
    Next i

    ' Clear original and rebuild from array
    For i = entries.count To 1 Step -1
        entries.Remove i
    Next i

    For i = 1 To UBound(entryArray)
        entries.Add entryArray(i)
    Next i

    Exit Sub

SortError:
    Debug.Print "[Sort Error] " & Err.Description
    MsgBox "Sorting failed: " & Err.Description, vbCritical
End Sub

' Inserts text into a table cell with a RichText content control
Public Sub InsertCellContent(cell As cell, content As String, tag As String, id As String)
    ' Existing logic
  '  Application.ScreenUpdating = False
    cell.Range.text = content
 '   Application.ScreenUpdating = True
    'Debug.Print "  InsertCellContent for tag '" & tag & "' took " & format(Timer - t0, "0.000") & " seconds."
End Sub

' User-callable test entry points
Public Sub PopulateBOMTable()
    Dim compounds As CompoundCollection
    Set compounds = GetGlobalCompoundCollection()
    Call PopulateCompoundsTable(compounds)
End Sub

Public Sub PopulateWSTable()
    Dim wastes As WasteCollection
    Set wastes = GetGlobalWasteCollection()
    Call PopulateWasteTable(wastes)
End Sub

'=== MODULE: IPCTablePopulator ===

' Populates the IPCs table from parsed unit operations
' Populates (or repopulates) the IPCs table, preserving user-entered remarks.
Public Sub PopulateIPCsTable()
    On Error GoTo ErrorHandler

    Dim ccIPCs       As ContentControl
    Dim ipcTable      As Table
    Dim row           As row
    Dim unitOps       As Collection
    Dim unitOp        As clsUnitOperation
    Dim i             As Long
    Dim titleLower    As String

    ' Dictionary to cache existing remarks by unitOp.id
    Dim remarkDict    As Object
    Set remarkDict = CreateObject("Scripting.Dictionary")

    Application.ScreenUpdating = False

    ' Locate IPCs table content control
    For Each ccIPCs In ThisDocument.contentControls
        If ccIPCs.title = "IPCs" Then
            If ccIPCs.Range.Tables.count > 0 Then
                Set ipcTable = ccIPCs.Range.Tables(1)
                Exit For
            End If
        End If
    Next ccIPCs

    If ipcTable Is Nothing Then
        MsgBox "Error: Could not find IPCs table.", vbExclamation
        GoTo Cleanup
    End If

    ' === 1) Cache existing remarks by unitOp.id (unlock step CCs so we can delete rows) ===
    Dim ccStep       As ContentControl
    Dim existingID   As String
    Dim remText      As String

    For Each row In ipcTable.rows
        If row.index > 1 Then   ' skip header row
            existingID = ""

            ' Read the CC in Column 1 (if it exists), capture its Title, then unlock it
            With row.Cells(1).Range
                If .contentControls.count > 0 Then
                    Set ccStep = .contentControls(1)
                    existingID = ccStep.title

                    ' Unlock the CC so deleting the row will be allowed
                    ccStep.LockContentControl = False
                    ccStep.LockContents = False
                End If
            End With

            If Len(existingID) > 0 Then
                ' Column 4: Remarks
                With row.Cells(4).Range
                    If .contentControls.count > 0 Then
                        remText = Trim(Replace(.contentControls(1).Range.text, Chr(13) & Chr(7), ""))
                    Else
                        remText = Trim(Replace(.text, Chr(13) & Chr(7), ""))
                    End If
                End With

                remarkDict(existingID) = remText
            End If
        End If
    Next row

    ' === 2) Clear all rows except the header row ===
    While ipcTable.rows.count > 1
        ipcTable.rows(2).Delete
    Wend

    ' === 3) Parse unit operations from the ProcessDescription ===
    Set unitOps = ParseProcessDescription()

    ' === 4) Rebuild table rows, restoring remarks if present ===
    For i = 1 To unitOps.count
        Set unitOp = unitOps(i)
        titleLower = LCase(Trim(unitOp.title))

        If titleLower = "ipc" Or titleLower = "pi" Then

            ' Add a new row
            Set row = ipcTable.rows.Add
            row.Range.Font.Bold = False

            ' --- Column 1: Step Number wrapped in a CC titled with unitOp.id ---
            Set ccStep = row.Cells(1).Range.contentControls.Add(wdContentControlRichText)
            With ccStep
                .title = unitOp.id                 ' store the unique ID
                .LockContentControl = True         ' prevent user from deleting the CC
                .LockContents = False              ' allow editing the step text
                .Range.text = CStr(unitOp.Step)    ' visible step number
            End With

            ' --- Column 2: IPC ID or PI ID ---
            Dim idValue As String
            If titleLower = "ipc" Then
                idValue = GetParamValue(unitOp, "IPC ID")
            Else
                idValue = GetParamValue(unitOp, "PI ID")
            End If

            Call InsertCellContent( _
                row.Cells(2), _
                idValue, _
                titleLower & "_id", _
                unitOp.id & "IPC" _
            )

            ' === Column 3: Details (three lines with placeholders & bold reactor) ===
            Dim testItemVal    As String
            Dim reactorVal     As String
            Dim testMethodVal  As String
            Dim acceptVal      As String
            Dim line1          As String
            Dim line2          As String
            Dim line3          As String
            Dim detailText     As String
            Dim rngDetail      As Range

            testItemVal = GetParamValue(unitOp, "IPC test item")
            reactorVal = GetParamValue(unitOp, "reactor")
            testMethodVal = GetParamValue(unitOp, "IPC test method")
            acceptVal = GetParamValue(unitOp, "acceptance criteria")

            ' Substitute "---" for any missing field
            If testItemVal = "" Then testItemVal = "---"
            If reactorVal = "" Then reactorVal = "---"
            If testMethodVal = "" Then testMethodVal = "---"
            If acceptVal = "" Then acceptVal = "---"

            ' Build the three-line block
            line1 = testItemVal & ", " & reactorVal
            line2 = testMethodVal & ":"
            line3 = acceptVal

            detailText = line1 & vbCrLf & line2 & vbCrLf & line3

            ' Insert into cell via helper
            Call InsertCellContent( _
                row.Cells(3), _
                detailText, _
                "details", _
                unitOp.id & "IPC" _
            )

            ' Apply bold to reactor portion if actual reactor is not "---"
            If reactorVal <> "---" Then
                Set rngDetail = row.Cells(3).Range
                rngDetail.End = rngDetail.End - 1   ' Exclude end-of-cell marker
                With rngDetail.Find
                    .ClearFormatting
                    .text = reactorVal
                    .MatchCase = True
                    .MatchWholeWord = True
                    If .Execute Then
                        rngDetail.Font.Bold = True
                    End If
                End With
            End If

            ' --- Column 4: Remarks (restore if cached) ---
            Dim cachedRemark As String
            If remarkDict.Exists(unitOp.id) Then
                cachedRemark = remarkDict(unitOp.id)
            Else
                cachedRemark = ""
            End If

            Call InsertCellContent( _
                row.Cells(4), _
                CStr(cachedRemark), _
                "remarks", _
                unitOp.id & "IPC" _
            )
        End If
    Next i

Cleanup:
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Debug.Print "Error in PopulateIPCsTable: " & Err.Description
    MsgBox "Failed to populate IPCs table: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

' Helper function to retrieve a parameter value (or return "---" if missing)
Private Function GetParamValue( _
    ByVal unitOp As clsUnitOperation, _
    ByVal paramName As String _
) As String
    Dim param As Object
    ' 1) Look in Parameters
    For Each param In unitOp.Parameters
        If LCase(Trim(param("Name"))) = LCase(Trim(paramName)) Then
            GetParamValue = param("Value")
            Exit Function
        End If
    Next param

    ' 2) If paramName = "reactor", look in Equipment
    If LCase(Trim(paramName)) = "reactor" Then
        For Each param In unitOp.Equipment
            GetParamValue = param("Text")
            Exit Function
        Next param
    End If

    ' 3) Not found ? return placeholder
    GetParamValue = "---"
End Function

'=== MODULE: HoldPointsTablePopulator ===

' Populates (or repopulates) the Hold Points table, preserving user-entered remarks.
' Populates (or repopulates) the Hold Points table, preserving user-entered remarks.
Public Sub PopulateHoldPointsTable()
    On Error GoTo ErrorHandler

    Dim ccHold        As ContentControl
    Dim holdTable     As Table
    Dim row           As row
    Dim unitOps       As Collection
    Dim unitOp        As clsUnitOperation
    Dim i             As Long
    Dim opTitle       As String

    ' Dictionary to cache existing remark by unitOp.id
    Dim remarkDict    As Object
    Set remarkDict = CreateObject("Scripting.Dictionary")

    Application.ScreenUpdating = False

    ' Locate Hold Points table content control
    For Each ccHold In ThisDocument.contentControls
        If ccHold.title = "Hold Points" Then
            If ccHold.Range.Tables.count > 0 Then
                Set holdTable = ccHold.Range.Tables(1)
                Exit For
            End If
        End If
    Next ccHold

    If holdTable Is Nothing Then
        MsgBox "Error: Could not find Hold Points table.", vbExclamation
        GoTo Cleanup
    End If

    ' === 1) Cache existing remark by unitOp.id (and unlock CCs so we can delete rows) ===
    Dim ccStep       As ContentControl
    Dim existingID   As String
    Dim remark       As String

    For Each row In holdTable.rows
        If row.index > 1 Then   ' skip header row
            existingID = ""

            ' Read the CC in Column 1 (if it exists), capture its Title, then unlock it
            With row.Cells(1).Range
                If .contentControls.count > 0 Then
                    Set ccStep = .contentControls(1)
                    existingID = ccStep.title

                    ' Unlock this CC so that deleting its row will be allowed
                    ccStep.LockContentControl = False
                    ccStep.LockContents = False
                End If
            End With

            If Len(existingID) > 0 Then
                ' Column 5: Remarks
                With row.Cells(5).Range
                    If .contentControls.count > 0 Then
                        remark = Trim(Replace(.contentControls(1).Range.text, Chr(13) & Chr(7), ""))
                    Else
                        remark = Trim(Replace(.text, Chr(13) & Chr(7), ""))
                    End If
                End With

                remarkDict(existingID) = remark
            End If
        End If
    Next row

    ' === 2) Clear all rows except the header row ===
    While holdTable.rows.count > 1
        holdTable.rows(2).Delete
    Wend

    ' === 3) Parse unit operations from the ProcessDescription ===
    Set unitOps = ParseProcessDescription()

    ' === 4) Rebuild table rows, restoring remarks if present ===
    For i = 1 To unitOps.count
        Set unitOp = unitOps(i)
        opTitle = LCase(Trim(unitOp.title))

        If opTitle = "hold point" Then

            ' Add a new row
            Set row = holdTable.rows.Add
            row.Range.Font.Bold = False

            ' --- Column 1: Step Number wrapped in a CC titled with unitOp.id ---
            Set ccStep = row.Cells(1).Range.contentControls.Add(wdContentControlRichText)
            With ccStep
                .title = unitOp.id                ' store the unique ID
                .LockContentControl = True         ' prevent user from deleting the CC
                .LockContents = False              ' allow editing the step text
                .Range.text = CStr(unitOp.Step)    ' visible step number
            End With

            ' --- Column 2: Reactor ---
            Dim reactorValue As String
            reactorValue = GetParamValue(unitOp, "reactor")
            Call InsertCellContent( _
                row.Cells(2), _
                reactorValue, _
                "reactor", _
                unitOp.id & "HP" _
            )

            ' --- Column 3: Time ---
            Dim timeValue As String
            timeValue = GetParamValue(unitOp, "time")
            Call InsertCellContent( _
                row.Cells(3), _
                timeValue, _
                "time", _
                unitOp.id & "HP" _
            )

            ' --- Column 4: Temperature ---
            Dim tempValue As String
            tempValue = GetParamValue(unitOp, "temperature")
            Call InsertCellContent( _
                row.Cells(4), _
                tempValue, _
                "temperature", _
                unitOp.id & "HP" _
            )

            ' --- Column 5: Remarks (restore if cached) ---
            Dim cachedRemark As String
            If remarkDict.Exists(unitOp.id) Then
                cachedRemark = remarkDict(unitOp.id)
            Else
                cachedRemark = ""
            End If

            Call InsertCellContent( _
                row.Cells(5), _
                CStr(cachedRemark), _
                "remarks", _
                unitOp.id & "HP" _
            )
        End If
    Next i

Cleanup:
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Debug.Print "Error in PopulateHoldPointsTable: " & Err.Description
    MsgBox "Failed to populate Hold Points table: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

'=== MODULE: ClearPointsTablePopulator ===

'=== MODULE: ClearPointsTablePopulator ===

' Populates the Clear Points table, wrapping each Step in a hidden CC titled with the unitOp.id
'=== MODULE: ClearPointsTablePopulator ===

' Populates (or repopulates) the Clear Points table, preserving user-entered remarks.
'=== MODULE: ClearPointsTablePopulator ===

' Populates (or repopulates) the Clear Points table, preserving user-entered remarks.
' Populates (or repopulates) the Clear Points table, preserving user-entered remarks.
' Populates (or repopulates) the Clear Points table, preserving user-entered remarks.
Public Sub PopulateClearPointsTable()
    On Error GoTo ErrorHandler

    Dim ccClear      As ContentControl
    Dim clearTable   As Table
    Dim row          As row
    Dim unitOps      As Collection
    Dim unitOp       As clsUnitOperation
    Dim i            As Long
    Dim opTitle      As String

    ' Dictionary to cache existing remarks by unitOp.id
    Dim remarkDict   As Object
    Set remarkDict = CreateObject("Scripting.Dictionary")

    Application.ScreenUpdating = False

    ' Locate Clear Points table content control
    For Each ccClear In ThisDocument.contentControls
        If ccClear.title = "Clear Points" Then
            If ccClear.Range.Tables.count > 0 Then
                Set clearTable = ccClear.Range.Tables(1)
                Exit For
            End If
        End If
    Next ccClear

    If clearTable Is Nothing Then
        MsgBox "Error: Could not find Clear Points table.", vbExclamation
        GoTo Cleanup
    End If

    ' === 1) Cache existing remarks by unitOp.id (and unlock CCs so we can delete rows) ===
    Dim ccStep       As ContentControl
    For Each row In clearTable.rows
        If row.index > 1 Then   ' skip header row
            Dim existingID As String
            existingID = ""

            ' Read the CC in Column 1 (if it exists), capture its Title, then unlock it
            With row.Cells(1).Range
                If .contentControls.count > 0 Then
                    Set ccStep = .contentControls(1)
                    existingID = ccStep.title

                    ' Unlock this CC so that deleting its row will be allowed
                    ccStep.LockContentControl = False
                    ccStep.LockContents = False
                End If
            End With

            If Len(existingID) > 0 Then
                Dim rem1 As String, rem2 As String

                ' Column 3: Remarks1
                With row.Cells(3).Range
                    If .contentControls.count > 0 Then
                        rem1 = Trim(Replace(.contentControls(1).Range.text, Chr(13) & Chr(7), ""))
                    Else
                        rem1 = Trim(Replace(.text, Chr(13) & Chr(7), ""))
                    End If
                End With

                ' Column 4: Remarks2
                With row.Cells(4).Range
                    If .contentControls.count > 0 Then
                        rem2 = Trim(Replace(.contentControls(1).Range.text, Chr(13) & Chr(7), ""))
                    Else
                        rem2 = Trim(Replace(.text, Chr(13) & Chr(7), ""))
                    End If
                End With

                remarkDict(existingID) = Array(rem1, rem2)
            End If
        End If
    Next row

    ' === 2) Clear all rows except the header row ===
    While clearTable.rows.count > 1
        clearTable.rows(2).Delete
    Wend

    ' === 3) Parse unit operations from the ProcessDescription ===
    Set unitOps = ParseProcessDescription()

    ' === 4) Rebuild table rows, restoring remarks if present ===
    For i = 1 To unitOps.count
        Set unitOp = unitOps(i)
        opTitle = LCase(Trim(unitOp.title))

        If opTitle = "transfer" _
           Or opTitle = "polish filter" _
           Or opTitle = "phase separation" _
           Or opTitle = "prepare solution" Then

            ' Add a new row
            Set row = clearTable.rows.Add
            row.Range.Font.Bold = False

            ' --- Column 1: Step Number wrapped in a CC titled with unitOp.id ---
            Set ccStep = row.Cells(1).Range.contentControls.Add(wdContentControlRichText)
            With ccStep
                .title = unitOp.id                ' store the unique ID
                .LockContentControl = True         ' prevent user from deleting the CC
                .LockContents = False              ' allow editing the step text
                .Range.text = CStr(unitOp.Step)    ' visible step number
            End With

            ' --- Column 2: Unit Operation Title ---
            Call InsertCellContent( _
                row.Cells(2), _
                unitOp.title, _
                "unit_op_title", _
                unitOp.id & "CP" _
            )

            ' Retrieve cached remarks (if any); otherwise use empty strings
            Dim cachedRems As Variant
            If remarkDict.Exists(unitOp.id) Then
                cachedRems = remarkDict(unitOp.id)
            Else
                cachedRems = Array("", "")
            End If

            ' --- Column 3: Remarks1 (restore if cached) ---
            Call InsertCellContent( _
                row.Cells(3), _
                CStr(cachedRems(0)), _
                "remarks1", _
                unitOp.id & "CP" _
            )

            ' --- Column 4: Remarks2 (restore if cached) ---
            Call InsertCellContent( _
                row.Cells(4), _
                CStr(cachedRems(1)), _
                "remarks2", _
                unitOp.id & "CP" _
            )
        End If
    Next i

Cleanup:
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Debug.Print "Error in PopulateClearPointsTable: " & Err.Description
    MsgBox "Failed to populate Clear Points table: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

' (Your existing GetParamValue helper can remain unchanged.)

'=== MODULE: BOMSummaryPopulator ===

Public Sub PopulateBOMSummaryTable()
    On Error GoTo ErrorHandler

    Dim t0 As Double, t1 As Double, t2 As Double, t3 As Double: t0 = Timer

    Dim cc As ContentControl, tbl As Table, rw As row
    Dim compounds As CompoundCollection
    Set compounds = GetGlobalCompoundCollection()

    Application.ScreenUpdating = False

    ' 1) Find "BOM Summary" table
    Set tbl = FindTableByTitle("BOM Summary")
    If tbl Is Nothing Then
        MsgBox "Error: Could not find 'BOM Summary' table.", vbExclamation
        GoTo Cleanup
    End If

    ' 2) Clear table (keep header)
    While tbl.rows.count > 1
        tbl.rows(2).Delete
    Wend
    t1 = Timer

' 3) Group by (code, name): TotalMass + TotalVolume + unique Steps + FirstStep
Dim buckets As Object: Set buckets = CreateObject("Scripting.Dictionary")
Dim i As Long
For i = 1 To compounds.count
    Dim cmp As Compound
    Set cmp = compounds.item(i)

    Dim code As String: code = Trim(CStr(cmp.Stoffdaten.productCode))
    If Len(code) = 0 Then code = "---"

    Dim name As String: name = Trim(CStr(cmp.Stoffdaten.title))
    If Len(name) = 0 Then name = "---"

    Dim key As String: key = code & vbNullChar & name

    Dim stepNo As Long: stepNo = CLng(GetStepNumberFromId(cmp.id))
    Dim mass As Double: mass = 0
    Dim volMl As Double: volMl = 0

    On Error Resume Next
    mass = CDbl(cmp.amount.mass)
    volMl = CDbl(cmp.amount.Volume)
    On Error GoTo ErrorHandler

    Dim b As Object
    If Not buckets.Exists(key) Then
        Set b = CreateObject("Scripting.Dictionary")
        b("Code") = code
        b("Name") = name
        b("TotalMass") = 0#
        b("TotalVolume") = 0#   ' NEW
        b("FirstStep") = stepNo
        Set b("Steps") = CreateObject("Scripting.Dictionary")
        buckets.Add key, b
    Else
        Set b = buckets(key)
        If stepNo < CLng(b("FirstStep")) Then b("FirstStep") = stepNo
    End If

    b("TotalMass") = CDbl(b("TotalMass")) + mass
    b("TotalVolume") = CDbl(b("TotalVolume")) + volMl   ' NEW

    If Not b("Steps").Exists(CStr(stepNo)) Then b("Steps").Add CStr(stepNo), True
Next i

    t2 = Timer

    ' 4) Copy buckets to array and sort (by Code, then Name)
    Dim arr() As Variant, idx As Long: idx = 0
    ReDim arr(1 To buckets.count)
    Dim k As Variant
    For Each k In buckets.keys
        idx = idx + 1
        Set arr(idx) = buckets(k)
    Next k

    Call SortSummaryEntries(arr) ' stable, simple sort by Code, then Name
    t3 = Timer

    ' 5) Render rows
    Dim j As Long
    For j = 1 To UBound(arr)
        Dim bb As Object: Set bb = arr(j)

        Dim stepsCsv As String: stepsCsv = JoinSortedSteps(bb("Steps"))
        Dim codeOut As String:  codeOut = CStr(bb("Code"))
        Dim nameOut As String:  nameOut = CStr(bb("Name"))
        Dim massOut As String:  massOut = SigStr(CDbl(bb("TotalMass")))
        Dim volOut  As String:  volOut = SigStr(CDbl(bb("TotalVolume")))

        Set rw = tbl.rows.Add
        rw.Range.Font.Bold = False

        ' Col 1: Steps; Col 2: Code; Col 3: Name; Col 4: Sum mass
        Call InsertCellContent(rw.Cells(1), stepsCsv, "steps", codeOut & nameOut & "SUM")
        Call InsertCellContent(rw.Cells(2), codeOut, "product_code", codeOut & nameOut & "SUM")
        Call InsertCellContent(rw.Cells(3), nameOut, "product_name", codeOut & nameOut & "SUM")
        Call InsertCellContent(rw.Cells(4), massOut, "sum_weight", codeOut & nameOut & "SUM")
        Call InsertCellContent(rw.Cells(5), volOut, "sum_volume", codeOut & nameOut & "SUM")
    Next j

    Debug.Print "TIMING (BOM Summary):"
    Debug.Print "  Clear: " & format(t1 - t0, "0.000") & " s"
    Debug.Print "  Group: " & format(t2 - t1, "0.000") & " s"
    Debug.Print "  Sort:  " & format(t3 - t2, "0.000") & " s"
    Debug.Print "  Total: " & format(Timer - t0, "0.000") & " s"

Cleanup:
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Debug.Print "Error in PopulateBOMSummaryTable: " & Err.Description
    MsgBox "Failed to populate BOM Summary: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

'--- Helpers ---

Private Function FindTableByTitle(ByVal titleText As String) As Table
    Dim cc As ContentControl
    For Each cc In ThisDocument.contentControls
        If cc.title = titleText Then
            If cc.Range.Tables.count > 0 Then
                Set FindTableByTitle = cc.Range.Tables(1)
                Exit Function
            End If
        End If
    Next cc
    Set FindTableByTitle = Nothing
End Function

Private Function JoinSortedSteps(ByVal stepSet As Object) As String
    Dim tmp() As Long, n As Long, k As Variant
    If stepSet Is Nothing Or stepSet.count = 0 Then
        JoinSortedSteps = ""
        Exit Function
    End If
    ReDim tmp(1 To stepSet.count)
    For Each k In stepSet.keys
        n = n + 1
        tmp(n) = CLng(k)
    Next k
    ' simple ascending sort
    Dim i As Long, j As Long, t As Long
    For i = 1 To n - 1
        For j = i + 1 To n
            If tmp(i) > tmp(j) Then
                t = tmp(i): tmp(i) = tmp(j): tmp(j) = t
            End If
        Next j
    Next i
    Dim s() As String: ReDim s(1 To n)
    For i = 1 To n
        s(i) = CStr(tmp(i))
    Next i
    JoinSortedSteps = Join(s, ", ")
End Function

Private Sub SortSummaryEntries(ByRef arr() As Variant)
    On Error GoTo SortError

    ' Guard: empty / 1-element arrays
    Dim ub As Long
    On Error Resume Next
    ub = UBound(arr)
    If Err.number <> 0 Or ub <= 1 Then Exit Sub
    On Error GoTo SortError

    Dim i As Long, j As Long
    For i = 1 To ub - 1
        For j = i + 1 To ub
            Dim a As Object, b As Object
            Set a = arr(i)
            Set b = arr(j)

            ' Sort by earliest step asc, then code, then name
            If (CLng(a("FirstStep")) > CLng(b("FirstStep"))) _
               Or (CLng(a("FirstStep")) = CLng(b("FirstStep")) And CStr(a("Code")) > CStr(b("Code"))) _
               Or (CLng(a("FirstStep")) = CLng(b("FirstStep")) And CStr(a("Code")) = CStr(b("Code")) And CStr(a("Name")) > CStr(b("Name"))) Then
                Dim tmp As Object
                Set tmp = arr(i)
                Set arr(i) = arr(j)
                Set arr(j) = tmp
            End If
        Next j
    Next i
    Exit Sub

SortError:
    Debug.Print "[Sort Error] " & Err.Description
End Sub





