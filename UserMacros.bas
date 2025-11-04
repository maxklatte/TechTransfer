Attribute VB_Name = "UserMacros"
'=== MODULE: UserMacros ===
Option Explicit

' Define Word constants for fallback

Const wdContentControlDropdown As Long = 3
Private internal_EnableTableUpdates As Variant
Private internal_EnableUserInteraction As Variant
Private internal_EventsEnabled As Variant


Public Sub ScaleUp()
    On Error GoTo HandleError
    Debug.Print "[INIT] Starting scale-up process for Compounds and Wastes."

    Dim factorInput As String
    factorInput = InputBox("Enter scale-up factor (e.g. 2 for double, 0.5 for half):", "Scale Up")

    If Not IsNumeric(factorInput) Or val(factorInput) <= 0 Then
        MsgBox "Invalid scale factor.", vbExclamation
        Exit Sub
    End If

    Dim factor As Double: factor = val(factorInput)

    Dim ccCount As Long: ccCount = GetGlobalCompoundCollection().count
    Dim wcCount As Long: wcCount = GetGlobalWasteCollection().count

    If ccCount = 0 And wcCount = 0 Then
        MsgBox "No compounds or wastes to scale.", vbInformation
        Exit Sub
    End If

GetGlobalCompoundCollection().ScaleAllMasses factor
GetGlobalWasteCollection().ScaleAll factor
Call RefreshCompoundDisplays
Call RefreshWasteDisplays
Call PopulateWSTable
Call PopulateBOMTable
Call PopulateBOMSummaryTable
Call PopulateMassBalanceTable


    Debug.Print "[DONE] ScaleUp complete."

    Exit Sub
HandleError:
    Debug.Print "[ERROR] ScaleUp: " & Err.Description
End Sub

' Move to UnitOperationManager => add wrapper frunction here
Sub AddDropdownToTableRow()
    On Error GoTo ErrorHandler
    ' Debug.Print "Starting AddDropdownToTableRow..."

    Dim targetTable As Table
    Dim sourceTable As Table
    Dim cursorRow As row
    Dim ccDropdown As ContentControl
    Dim unitOperationRow As row
    Dim ccUnitOperation As ContentControl
    Dim dropdownDisplayName As String
    Dim dropdownKey As String
    Dim dropdownEntry As ContentControlListEntry


    ' Get the source and target tables via their content controls
    Set sourceTable = GetTableFromContentControl("DefinitionOfUnitOperations")
    Set targetTable = GetTableFromContentControl("ProcessDescription")

    ' Validate that both tables are available
    If sourceTable Is Nothing Then
        Debug.Print "Error: Source table (DefinitionOfUnitOperations) not found."
        Exit Sub
    Else
        ' Debug.Print "Source table (DefinitionOfUnitOperations) found."
    End If

    If targetTable Is Nothing Then
        Debug.Print "Error: Target table (ProcessDescription) not found."
        Exit Sub
    Else
        ' Debug.Print "Target table (ProcessDescription) found."
    End If

    ' Ensure the cursor is in the target table
    If Not Selection.Information(wdWithInTable) Then
        Debug.Print "Error: Cursor is not in any table."
        Exit Sub
    End If

    ' Check if the selected table is the target table
    If Not UnitOperationManager.AreTablesEqual(Selection.Tables(1), targetTable) Then
        Debug.Print "Error: Cursor is not in the ProcessDescription table."
        Exit Sub
    Else
        'Debug.Print "Cursor confirmed in ProcessDescription table."
    End If

    ' Get the current row where the cursor is located
    Set cursorRow = Selection.rows(1)
   ' Debug.Print "Cursor located at row index: " & cursorRow.index

    ' Add a dropdown content control to the third column of the current row
    Set ccDropdown = cursorRow.Cells(2).Range.contentControls.Add(wdContentControlDropdown)
   ' Debug.Print "Dropdown content control added to the third column."

    ' Populate dropdown with entries from the source table
    For Each unitOperationRow In sourceTable.rows
        ' Extract the content control from the first column
        Set ccUnitOperation = FetchContentControlFromCell(unitOperationRow.Cells(1))
        If Not ccUnitOperation Is Nothing Then
            dropdownKey = Trim(ccUnitOperation.title) ' Key for the dropdown (used in HandleDropdownSelection)
            dropdownDisplayName = Trim(ccUnitOperation.Range.text) & "-" & dropdownKey ' Display name
            If dropdownKey <> "" And dropdownDisplayName <> "" Then
                ' Add the display name to the dropdown
                Set dropdownEntry = ccDropdown.DropdownListEntries.Add(dropdownDisplayName)
                ' Store the key in the Value property
                dropdownEntry.value = dropdownKey
               ' Debug.Print "Added dropdown entry: DisplayName = " & dropdownDisplayName & ", Key = " & dropdownKey
            Else
                Debug.Print "Skipped row with empty display name or key."
            End If
        Else
            Debug.Print "No content control found in the first cell of row " & unitOperationRow.index
        End If
    Next unitOperationRow

    ' Configure the dropdown
    ccDropdown.SetPlaceholderText text:="Select a unit operation..."
    ccDropdown.title = "UnitOperationDropdown"
    ccDropdown.tag = "UnitOperationDropdown"
    ' Debug.Print "Dropdown configuration completed successfully."


    Exit Sub

ErrorHandler:
    Debug.Print "Error in AddDropdownToTableRow: " & Err.Description
    Err.Clear
End Sub


Public Sub SelectReferenceCompound()
    Dim selected As Compound
    Set selected = GetGlobalCompoundCollection().PromptCompoundSelection()
    If Not selected Is Nothing Then
        GetGlobalCompoundCollection().ReferenceCompoundId = selected.id
        Debug.Print "[INFO] Reference compound set to: " & selected.id
    End If
    ' Update all input fields and relevant tables
    ContentControlHelpers.RefreshInputFieldsLow
    PopulateBOMTable
    PopulateBOMSummaryTable
    PopulateMassBalanceTable
End Sub

'ForceReloadStoffdatenbank
'UserInteractionToggle
'AutoUpdateTablesToggle
'GenerateFlowChart
'Laborvorschrift

Public Sub ForceReloadStoffdatenbank()
    Const filePath As String = "https://cordenpharma.sharepoint.com/sites/cp_sui_intranetsui/prd/Freigegebene Dokumente/FE Stoffdatenbank.xlsm"
    Const sheetName As String = "Stoffdatenbank"

    Dim userConfirm As VbMsgBoxResult
    userConfirm = MsgBox("This will forcibly reload the Stoffdatenbank from SharePoint." & vbCrLf & _
                         "This may take a minute or more. Continue?", vbYesNo + vbExclamation, "Force Reload Confirmation")
    If userConfirm <> vbYes Then
        Debug.Print "[INFO] User cancelled force reload."
        Exit Sub
    End If

    On Error GoTo ReloadFail
    
    Set gStoffdatenbank = New Stoffdatenbank

    Dim startTime As Single: startTime = Timer
    gStoffdatenbank.LoadStoffeFromExcel filePath, sheetName
    Dim elapsed As Single: elapsed = Timer - startTime

    MsgBox "Stoffdatenbank forcibly reloaded in " & format(elapsed, "0.00") & " seconds.", vbInformation, "Reload Complete"
    Debug.Print "[INFO] Stoffdatenbank forcibly reloaded in " & format(elapsed, "0.00") & " seconds."
    Exit Sub

ReloadFail:
    MsgBox "[ERROR] Failed to forcibly reload Stoffdatenbank: " & Err.Description, vbCritical, "Reload Failed"
    Debug.Print "[ERROR] Reload failure: " & Err.Description
End Sub

Public Function IsTableUpdateEnabled() As Boolean
    If isEmpty(internal_EnableTableUpdates) Then
        internal_EnableTableUpdates = False
    End If
    IsTableUpdateEnabled = internal_EnableTableUpdates
End Function

Public Sub ToggleTableUpdateState()
    If IsTableUpdateEnabled() Then
        internal_EnableTableUpdates = False
    Else
        internal_EnableTableUpdates = True
    End If
    MsgBox "Table updates are now " & IIf(internal_EnableTableUpdates, "ENABLED", "DISABLED")
End Sub

Public Function IsUserInteractionEnabled() As Boolean
    If isEmpty(internal_EnableUserInteraction) Then
        internal_EnableUserInteraction = False
    End If
    IsUserInteractionEnabled = internal_EnableUserInteraction
End Function

Public Sub ToggleUserInteraction()
    internal_EnableUserInteraction = Not IsUserInteractionEnabled()
    MsgBox "User interaction macros are now " & IIf(IsUserInteractionEnabled(), "ENABLED", "DISABLED")
End Sub

Public Function IsEventsEnabled() As Boolean
    If isEmpty(internal_EventsEnabled) Then
        internal_EventsEnabled = True     ' default ON
    End If
    IsEventsEnabled = internal_EventsEnabled
End Function

' --- Control toggles ---
Public Sub DisableEvents()
    internal_EventsEnabled = False
End Sub

Public Sub EnableEvents()
    internal_EventsEnabled = True
End Sub


' === USER DIALOG ENTRY POINT ===
Public Sub CloneUnitOpsDialog()
    Dim doc As Document: Set doc = ThisDocument
    Dim wasTracked As Boolean
    Dim didProcess As Boolean: didProcess = False
    Dim units As Collection

    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    wasTracked = doc.TrackRevisions
    doc.TrackRevisions = False

    ' === Parse document ===
    Set units = ParseProcessDescription()
    If units Is Nothing Or units.count < 1 Then
        MsgBox "Document must contain at least 1 unit operation.", vbExclamation, "No Unit Operations"
        GoTo Cleanup
    End If

    ' === Step 1: Ask clone range ===
    Dim rangeInput As String
    rangeInput = InputBox( _
        "Enter the step numbers of the unit operations to be cloned." & vbCrLf & _
        "Example: 2-4 will clone steps 2 through 4 inclusive.", _
        "Clone Unit Operations")
    If rangeInput = "" Then GoTo Cleanup

    Dim dashPos As Long
    Dim startIndex As Long, endIndex As Long
    dashPos = InStr(rangeInput, "-")
    If dashPos > 0 Then
        startIndex = CLng(Trim(Left(rangeInput, dashPos - 1)))
        endIndex = CLng(Trim(Mid(rangeInput, dashPos + 1)))
    Else
        startIndex = CLng(Trim(rangeInput))
        endIndex = startIndex
    End If

    ' === Step 2: Ask insertion point ===
    Dim insertAfterInput As String
    Dim insertAfterIndex As Long
    insertAfterInput = InputBox( _
        "After which step should the clones be inserted?" & vbCrLf & _
        "Example: 3 means the clones will appear after step 3.", _
        "Insert Position")
    If insertAfterInput = "" Then GoTo Cleanup
    insertAfterIndex = CLng(Trim(insertAfterInput))

    ' === Step 3: Validate indices ===
    If startIndex < 1 Or endIndex > units.count Or insertAfterIndex > units.count Or endIndex < startIndex Then
        MsgBox "Invalid index range. Please try again.", vbExclamation, "Input Error"
        GoTo Cleanup
    End If

    ' === Step 4: Show preview ===
    Dim previewLines As Collection
    Set previewLines = GetClonePreview(units, startIndex, endIndex, insertAfterIndex)

    Dim previewText As String
    previewText = "Preview of result:" & vbCrLf & String(40, "-") & vbCrLf

    Dim line As Variant
    For Each line In previewLines
        previewText = previewText & line & vbCrLf
    Next line

    Dim result As VbMsgBoxResult
    result = MsgBox(previewText & vbCrLf & "Proceed with cloning?", vbYesNoCancel + vbQuestion, "Confirm Cloning")

    Select Case result
        Case vbYes
            Call ExecuteCloningWorkflow(startIndex, endIndex, insertAfterIndex)
            Call Renumber
            didProcess = True
        Case vbNo
            MsgBox "Cloning canceled. Please re-enter your selection.", vbInformation
        Case vbCancel
            MsgBox "Operation canceled.", vbInformation
    End Select

    If didProcess Then doc.UndoClear

Cleanup:
    doc.TrackRevisions = wasTracked
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Dim logMsg As String
    logMsg = "[ERROR] " & format(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf & _
             "Macro:    CloneUnitOpsDialog" & vbCrLf & _
             "Doc:      " & ThisDocument.FullName & vbCrLf & _
             "User:     " & Environ("USERNAME") & vbCrLf & _
             "Err #:    " & Err.number & vbCrLf & _
             "Source:   " & Err.Source & vbCrLf & _
             "Message:  " & Err.Description

    Debug.Print logMsg
    MsgBox logMsg, vbExclamation, "Unexpected Error"
    Resume Cleanup
End Sub

Public Sub AddUnitOperation()
    Dim cc As ContentControl

    Set cc = ContentControlHelpers.ReturnSelectedContentControl()

    If cc Is Nothing Then
        MsgBox "No content control found at the current selection.", vbExclamation, "No Selection"
        Exit Sub
    End If

    If cc.title = "UnitOperationDropdown" Then
        Call UnitOperationManager.HandleDropdownSelection(cc)
    Else
        MsgBox "Selected content control is not a UnitOperationDropdown.", vbInformation, "Invalid Selection"
    End If
End Sub

Public Sub ExpandField()
    Dim cc As ContentControl
    Dim doc As Document
    Dim wasTracked As Boolean

    On Error GoTo ErrorHandler
    Set doc = ActiveDocument
    wasTracked = doc.TrackRevisions
    doc.TrackRevisions = False

    Set cc = ContentControlHelpers.ReturnSelectedContentControl()
    If cc Is Nothing Then
        MsgBox "Please place the cursor inside a field to expand.", vbInformation
        GoTo Cleanup
    End If

    Call ExpandContentControl(cc, doc)

Cleanup:
    doc.TrackRevisions = wasTracked
    Exit Sub

ErrorHandler:
    MsgBox "Error in ExpandField: " & Err.Description, vbCritical
    Resume Cleanup
End Sub


Public Sub ComputeField()
    On Error GoTo HandleError

    Dim cc As ContentControl
    Dim doc As Document
    Dim cmp As Compound
    Dim ref As Compound
    Dim w As Waste

    Set doc = ActiveDocument
    Set cc = ContentControlHelpers.ReturnSelectedContentControl()
    If cc Is Nothing Then
        MsgBox "Please place the cursor inside a supported field.", vbExclamation
        Exit Sub
    End If

    If Not ContentControlHelpers.ParentContentControlTitled("ProcessDescription", cc) Then
        MsgBox "Selected field is not inside a ProcessDescription block.", vbInformation
        Exit Sub
    End If

    Select Case LCase(cc.title)
        Case "input", "product"
            Set cmp = GetGlobalCompoundCollection().GetById(cc.tag)
            If cmp Is Nothing Then
                MsgBox "No compound found for this field.", vbCritical
                Exit Sub
            End If

            Call ComputeCompoundEdit(cc)
            Set ref = GetGlobalCompoundCollection().GetReferenceCompound()

            Application.ScreenUpdating = False
            cc.Range.text = cmp.ToDisplayString("edit", ref)
            cc.Range.Font.Bold = True
            Application.ScreenUpdating = True

        Case "waste"
            Set w = GetGlobalWasteCollection().GetById(cc.tag)
            If w Is Nothing Then
                MsgBox "No waste found for this field.", vbCritical
                Exit Sub
            End If

            Call ComputeWasteEdit(cc)

            Application.ScreenUpdating = False
            cc.Range.text = w.ToDisplayString("edit")
            cc.Range.Font.Bold = True
            Application.ScreenUpdating = True

        Case Else
            MsgBox "ComputeField does not yet support this field type (" & cc.title & ").", vbInformation
    End Select

    Exit Sub

HandleError:
    MsgBox "Error in ComputeField: " & Err.Description, vbCritical
    Err.Clear
End Sub

'=== NEW: Refresh Stoffdaten for selected field (input/product only) ===
Public Sub RefreshStoffdatenForField()
    On Error GoTo HandleError

    Dim cc As ContentControl
    Dim doc As Document
    Dim cmp As Compound
    Dim ref As Compound

    ' Identify active ContentControl
    Set doc = ActiveDocument
    Set cc = ContentControlHelpers.ReturnSelectedContentControl()
    If cc Is Nothing Then
        MsgBox "Please place the cursor inside a supported field.", vbExclamation
        Exit Sub
    End If

    ' Ensure within a ProcessDescription block
    If Not ContentControlHelpers.ParentContentControlTitled("ProcessDescription", cc) Then
        MsgBox "Selected field is not inside a ProcessDescription block.", vbInformation
        Exit Sub
    End If

    Select Case LCase(cc.title)
        Case "input", "product"
            ' Lookup compound
            Set cmp = GetGlobalCompoundCollection().GetById(cc.tag)
            If cmp Is Nothing Then
                MsgBox "No compound found for this field.", vbCritical
                Exit Sub
            End If

            ' Perform Stoffdaten refresh
            Call CompoundMutator.RefreshStoffdaten(cmp)

            ' Update display
            Set ref = GetGlobalCompoundCollection().GetReferenceCompound()
            Application.ScreenUpdating = False
            cc.Range.text = cmp.ToDisplayString("edit", ref)
            cc.Range.Font.Bold = True
            Application.ScreenUpdating = True

        Case Else
            MsgBox "RefreshStoffdatenForField supports only 'input' or 'product' fields.", vbExclamation
    End Select

    Exit Sub

HandleError:
    MsgBox "Error in RefreshStoffdatenForField: " & Err.Description, vbCritical
    Err.Clear
End Sub

Public Sub CollapseField()
    Dim cc As ContentControl
    Dim doc As Document
    Dim wasTracked As Boolean

    On Error GoTo ErrorHandler
    Set doc = ActiveDocument
    wasTracked = doc.TrackRevisions
    doc.TrackRevisions = False

    Set cc = ContentControlHelpers.ReturnSelectedContentControl()
    If cc Is Nothing Then
        MsgBox "Please place the cursor inside a field to collapse.", vbInformation
        GoTo Cleanup
    End If

    Call CollapseContentControl(cc, doc)

Cleanup:
    doc.TrackRevisions = wasTracked
    Exit Sub

ErrorHandler:
    MsgBox "Error in CollapseField: " & Err.Description, vbCritical
    Resume Cleanup
End Sub




