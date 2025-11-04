Attribute VB_Name = "FlowChartCreator"
' === Module: FlowChartCreator ===
Option Explicit

' === Modified: GenerateFlowChart ===
Public Sub GenerateFlowChart()
    On Error GoTo ErrorHandler

    Dim tStart As Double
    tStart = Timer
    Application.ScreenUpdating = False

    Debug.Print "[GenerateFlowChart] Running full flow chart build sequence..."

    Dim unitOperations As Collection
    Set unitOperations = ParseProcessDescription()

    If unitOperations Is Nothing Or unitOperations.count = 0 Then
        MsgBox "No unit operations parsed. Flow chart cannot be generated.", vbExclamation, "No Data"
        Debug.Print "[GenerateFlowChart] No unit operations parsed. Aborting."
        GoTo CleanExit
    End If

    Dim templateTitle As String
    Dim awarenessSuffix As String
    Dim userChoice As VbMsgBoxResult

    userChoice = MsgBox("Use 'Simple Flow Chart Template'?" & vbCrLf & _
                        "(Click No to use 'Flow Chart Template')", vbYesNoCancel + vbQuestion, "Select Template")

    Select Case userChoice
        Case vbYes
            templateTitle = "Simple Flow Chart Template"
            awarenessSuffix = "weight_name" 'to be renamed to reference input reporting style
        Case vbNo
            templateTitle = "Flow Chart Template"
            awarenessSuffix = "high"       'to be renamed to reference input reporting style
        Case Else
            Debug.Print "[GenerateFlowChart] User canceled template selection."
            GoTo CleanExit
    End Select

    Debug.Print "[GenerateFlowChart] Using template: " & templateTitle
    Call BuildFlowChartFromUnitOperations(unitOperations, templateTitle, awarenessSuffix)
    Debug.Print "[GenerateFlowChart] Flow chart generation finished."

CleanExit:
    Debug.Print "[Timing] GenerateFlowChart Duration: " & format(Timer - tStart, "0.000") & " sec"
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Debug.Print "[GenerateFlowChart] Error: " & Err.Description
    MsgBox "An error occurred while generating the flow chart: " & Err.Description, vbCritical, "Flow Chart Error"
    Err.Clear
    GoTo CleanExit
End Sub


' === Modified: BuildFlowChartFromUnitOperations ===
Public Sub BuildFlowChartFromUnitOperations(ByVal unitOps As Collection, ByVal templateTitle As String, ByVal awarenessSuffix As String)
    On Error GoTo ErrorHandler
    Debug.Print "[FlowChartCreator] Starting flow chart generation..."

    Dim ccTemplates As ContentControl
    Dim ccFlowChart As ContentControl
    Dim ccTemplateBlock As ContentControl
    Dim unitOp As clsUnitOperation
    Dim foundTemplate As Boolean
    Dim insertPoint As Range
    Dim i As Long
    Dim newBlock As ContentControl

    Set ccTemplates = FindContentControlByTitle(templateTitle)
    If ccTemplates Is Nothing Then
        MsgBox "Error: '" & templateTitle & "' content control not found.", vbCritical, "Flow Chart Error"
        Debug.Print "[FlowChartCreator] Error: Template zone not found: " & templateTitle
        Exit Sub
    End If

    Set ccFlowChart = FindContentControlByTitle("Flow Chart")
    If ccFlowChart Is Nothing Then
        MsgBox "Error: 'Flow Chart' content control not found.", vbCritical, "Flow Chart Error"
        Debug.Print "[FlowChartCreator] Error: Output zone not found."
        Exit Sub
    End If

    On Error Resume Next
    ccFlowChart.Range.Delete
    If Err.number <> 0 Then
        Debug.Print "[FlowChartCreator] Warning: Unable to clear Flow Chart range - " & Err.Description
        Err.Clear
    Else
        Debug.Print "[FlowChartCreator] Cleared existing content in 'Flow Chart'."
    End If
    On Error GoTo ErrorHandler

    Set insertPoint = ccFlowChart.Range.Duplicate
    insertPoint.Collapse Direction:=wdCollapseEnd

    For i = unitOps.count To 1 Step -1
        Set unitOp = unitOps(i)
        foundTemplate = False

        For Each ccTemplateBlock In ccTemplates.Range.contentControls
            If ccTemplateBlock.title = unitOp.title Then
                ccTemplateBlock.Range.Copy
                insertPoint.PasteAndFormat wdFormatOriginalFormatting
                insertPoint.SetRange insertPoint.End, insertPoint.End

                Set newBlock = FindNewlyInsertedBlock(ccFlowChart, unitOp.title)
                If Not newBlock Is Nothing Then
                    newBlock.tag = unitOp.id
                    Call PopulateFlowChartBlock(unitOp, awarenessSuffix)
                Else
                    Debug.Print "[FlowChartCreator] Warning: Inserted block not found for: " & unitOp.title
                End If

                foundTemplate = True
                Exit For
            End If
        Next ccTemplateBlock

        If Not foundTemplate Then
            MsgBox "Template block for '" & unitOp.title & "' not found.", vbExclamation, "Missing Template"
            Debug.Print "[FlowChartCreator] Warning: Template block not found for unit operation titled '" & unitOp.title & "'."
        End If
    Next i

    Debug.Print "[FlowChartCreator] Flow chart generation completed."
    Call FinalizeFlowChartLayout
    Exit Sub

ErrorHandler:
    Debug.Print "[FlowChartCreator] Error: " & Err.Description
    MsgBox "An unexpected error occurred while generating the flow chart: " & Err.Description, vbCritical, "Flow Chart Error"
    Err.Clear
End Sub

Private Function FindNewlyInsertedBlock(ByVal container As ContentControl, ByVal title As String) As ContentControl
    Dim cc As ContentControl
    For Each cc In container.Range.contentControls
        If cc.title = title Then
            Set FindNewlyInsertedBlock = cc
            Exit Function
        End If
    Next cc
    Set FindNewlyInsertedBlock = Nothing
End Function

Private Sub PopulateFlowChartBlock(unitOp As clsUnitOperation, Optional defaultAwareness As String = "flowchart")
    On Error GoTo ErrorHandler

    Dim flowChartCC As ContentControl
    Dim nestedCC As ContentControl
    Dim fullTag As String
    Dim valueText As String
    Dim wasFound As Boolean
    Dim foundBlock As Boolean

    If unitOp Is Nothing Then
        Debug.Print "[FlowChartCreator] Error: unitOp is Nothing"
        Exit Sub
    End If
    If Trim(unitOp.id) = "" Then
        Debug.Print "[FlowChartCreator] Error: unitOp.ID is empty"
        Exit Sub
    End If

    foundBlock = False

    For Each flowChartCC In ThisDocument.contentControls
        If flowChartCC.tag = unitOp.id Then
            For Each nestedCC In flowChartCC.Range.contentControls
                Dim isCompound As Boolean
                isCompound = False

                Dim tagIndex As String
                tagIndex = Mid(nestedCC.tag, 6, 2)

                Dim item As Object
                For Each item In unitOp.Inputs
                    If Mid(item("Tag"), 6, 2) = tagIndex Then
                        isCompound = True
                        Exit For
                    End If
                Next item

                fullTag = NormalizeContentControlTag(nestedCC.tag, unitOp.id, isCompound, defaultAwareness)
                valueText = unitOp.GetTextByTag(fullTag, wasFound)

                If wasFound Then
                    nestedCC.tag = fullTag
                    nestedCC.Range.text = valueText
                Else
                    nestedCC.Range.text = "N/A"
                End If
            Next nestedCC
            foundBlock = True
            Exit For
        End If
    Next flowChartCC

    If Not foundBlock Then
        Debug.Print "[FlowChartCreator] Error: Could not locate block with Tag = " & unitOp.id & " to populate."
    End If

    Exit Sub

ErrorHandler:
    Debug.Print "[FlowChartCreator] ERROR #" & Err.number & " at line " & Erl & ": " & Err.Description
    Err.Clear
End Sub

Public Sub FinalizeFlowChartLayout()
    On Error GoTo ErrorHandler
    Debug.Print "[FlowChartCreator] Starting flow chart finalization..."

    Dim tStart As Double
    tStart = Timer
    Application.ScreenUpdating = False ' ? Boost performance

    Dim ccFlowChart As ContentControl
    Dim ccBlock As ContentControl
    Dim tbl As Table
    Dim para As Paragraph
    Dim firstParaStart As Long, lastParaEnd As Long

    Set ccFlowChart = FindContentControlByTitle("Flow Chart")
    If ccFlowChart Is Nothing Then
        Debug.Print "[FlowChartCreator] Error: 'Flow Chart' content control not found."
        GoTo CleanExit
    End If

    firstParaStart = ccFlowChart.Range.Paragraphs(1).Range.Start
    lastParaEnd = ccFlowChart.Range.Paragraphs(ccFlowChart.Range.Paragraphs.count).Range.End

    For Each ccBlock In ccFlowChart.Range.contentControls
        If Len(ccBlock.tag) = 5 And IsNumeric(ccBlock.tag) Then
            On Error Resume Next
            Set tbl = ccBlock.Range.Tables(1)
            If Not tbl Is Nothing And tbl.rows.count >= 2 Then
                tbl.rows(tbl.rows.count).Delete
                tbl.rows(1).Delete
                'Debug.Print "[FlowChartCreator] Trimmed table rows in block: " & ccBlock.title
            End If
            On Error GoTo 0
        End If
    Next ccBlock

    Call FlattenMiddleLayerContentControls(ccFlowChart)
    Debug.Print "[FlowChartCreator] Flattened all middle-layer controls in Flow Chart."

    For Each para In ccFlowChart.Range.Paragraphs
        If para.Range.Tables.count = 0 And Len(Trim(para.Range.text)) > 0 Then
            If para.Range.Start > firstParaStart And para.Range.End < lastParaEnd Then
                para.Range.Delete
                ' Debug.Print "[FlowChartCreator] Removed line break or non-table paragraph."
            End If
        End If
    Next para

    Debug.Print "[FlowChartCreator] Flow chart finalization completed."

CleanExit:
Debug.Print "[Timing] FinalizeFlowChartLayout Duration: " & format(Timer - tStart, "0.000") & " sec"
    Application.ScreenUpdating = True ' ? Restore state
    Exit Sub

ErrorHandler:
    Debug.Print "[FlowChartCreator] Error in FinalizeFlowChartLayout: " & Err.Description
    MsgBox "An error occurred during flow chart finalization: " & Err.Description, vbCritical, "Flow Chart Error"
    Err.Clear
    Resume CleanExit
End Sub


' This logic should be incoperated into its own class managing all aspects of ID, properties and recognition - this i getting out of hands :)
Private Function NormalizeContentControlTag(rawTag As String, unitOpId As String, Optional isCompound As Boolean = False, Optional defaultAwareness As String = "flowchart") As String
    If Left(rawTag, 5) <> "00000" Then
        NormalizeContentControlTag = rawTag
        Exit Function
    End If

    Dim basePart As String
    Dim suffix As String

    If InStr(rawTag, "-") > 0 Then
        basePart = Split(rawTag, "-")(0)
        suffix = "-" & Split(rawTag, "-")(1)
    ElseIf isCompound Then
        basePart = rawTag
        suffix = "-" & defaultAwareness
    Else
        basePart = rawTag
        suffix = ""
    End If

    basePart = unitOpId & Mid(basePart, 6)
    NormalizeContentControlTag = basePart & suffix
End Function


Public Sub RefreshFlowChart()
    On Error GoTo FatalError

    Dim t0 As Double, tParse As Double, tWalk As Double
    t0 = Timer

    Dim doc As Document: Set doc = ThisDocument
    Dim prevSU As Boolean: prevSU = Application.ScreenUpdating
    Dim prevTrack As Boolean: prevTrack = doc.TrackRevisions
    Dim prevAlerts As WdAlertLevel: prevAlerts = Application.DisplayAlerts

    Debug.Print String(60, "-")
    Debug.Print "[RefreshFlowChart] Begin refresh (flattened Flow Chart)."

    ' Guards
    Application.ScreenUpdating = False
    Application.DisplayAlerts = wdAlertsNone
    DisableEvents
    doc.TrackRevisions = False

    ' Locate Flow Chart
    Dim ccFlow As ContentControl
    Set ccFlow = FindContentControlByTitle("Flow Chart")
    If ccFlow Is Nothing Then
        Debug.Print "[RefreshFlowChart] ERROR: 'Flow Chart' content control not found."
        MsgBox "'Flow Chart' area not found.", vbExclamation, "Refresh Flow Chart"
        GoTo Cleanup
    End If

    ' Parse source
    Dim unitOps As Collection
    Set unitOps = ParseProcessDescription()
    If unitOps Is Nothing Or unitOps.count = 0 Then
        Debug.Print "[RefreshFlowChart] WARNING: ParseProcessDescription returned no unit operations."
        MsgBox "No unit operations parsed. Nothing to refresh.", vbExclamation, "Refresh Flow Chart"
        GoTo Cleanup
    End If

    tParse = Timer
    Debug.Print "[RefreshFlowChart] Parsed unit operations in " & format(tParse - t0, "0.000") & " s"
    Debug.Print "[RefreshFlowChart] UnitOps count = " & unitOps.count

    ' Walk direct children CCs of Flow Chart
    Dim total As Long, updated As Long, naNotFound As Long
    Dim naMalformed As Long, naMissingUO As Long, perItemErrs As Long

    Dim cc As ContentControl
    For Each cc In ccFlow.Range.contentControls
        total = total + 1
        On Error GoTo ItemError

        Dim tg As String: tg = Trim$(cc.tag)
        If Len(tg) < 7 Then
            ' Need at least 7 chars (5 for unit op id, + 2 for field id)
            cc.Range.text = "N/A"
            naMalformed = naMalformed + 1
            GoTo NextCC
        End If

        ' First 5 chars: unit operation id
        Dim unitOpId As String
        unitOpId = Left$(tg, 5) ' Robust enough per your note; no numeric check required

        ' Find matching unitOp by iterating (no dictionaries)
        Dim u As clsUnitOperation
        Dim foundUO As Boolean: foundUO = False
        For Each u In unitOps
            If StrComp(u.id, unitOpId, vbBinaryCompare) = 0 Then
                foundUO = True
                Exit For
            End If
        Next u

        If Not foundUO Or u Is Nothing Then
            cc.Range.text = "N/A"
            naMissingUO = naMissingUO + 1
            GoTo NextCC
        End If

        ' Ask the model for the display text; pass FULL tag (incl. input suffix if present)
        Dim wasFound As Boolean
        Dim valueText As String
        valueText = u.GetTextByTag(tg, wasFound)

        If wasFound Then
            cc.Range.text = valueText
            updated = updated + 1
        Else
            cc.Range.text = "N/A"
            naNotFound = naNotFound + 1
        End If

        GoTo NextCC

ItemError:
        perItemErrs = perItemErrs + 1
        Debug.Print "[RefreshFlowChart][ItemError] Title='" & cc.title & "' Tag='" & cc.tag & _
                    "' :: " & Err.number & " - " & Err.Description
        Err.Clear
        On Error GoTo FatalError
        cc.Range.text = "N/A" ' fail-safe per your policy

NextCC:
        ' continue
    Next cc

    tWalk = Timer
    Debug.Print "[RefreshFlowChart] Walked " & total & " CC(s) in " & format(tWalk - tParse, "0.000") & " s"
    Debug.Print "  Updated:       " & updated
    Debug.Print "  N/A (not found in model): " & naNotFound
    Debug.Print "  N/A (malformed tag):      " & naMalformed
    Debug.Print "  N/A (missing unit op):    " & naMissingUO
    Debug.Print "  Item errors:              " & perItemErrs
    Debug.Print "[RefreshFlowChart] TOTAL: " & format(Timer - t0, "0.000") & " s"
    Debug.Print String(60, "-")

Cleanup:
    On Error Resume Next
    doc.TrackRevisions = prevTrack
    EnableEvents
    Application.DisplayAlerts = prevAlerts
    Application.ScreenUpdating = prevSU
    On Error GoTo 0
    Exit Sub

FatalError:
    Debug.Print "[RefreshFlowChart][FATAL] " & Err.number & " - " & Err.Description
    MsgBox "Error during Flow Chart refresh: " & Err.Description, vbCritical, "Flow Chart Refresh"
    Resume Cleanup
End Sub




