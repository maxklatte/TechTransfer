Attribute VB_Name = "MBRCreator"
' === Module: FlowChartCreator ===
Option Explicit

' === Modified: GenerateFlowChart ===
Public Sub GenerateMBR()
    On Error GoTo ErrorHandler

    Dim tStart As Double
    tStart = Timer
    Application.ScreenUpdating = False

    Debug.Print "[GenerateMBR] Running full flow chart build sequence..."

    Dim unitOperations As Collection
    Set unitOperations = ParseProcessDescription()

    If unitOperations Is Nothing Or unitOperations.count = 0 Then
        MsgBox "No unit operations parsed. MBR cannot be generated.", vbExclamation, "No Data"
        Debug.Print "[GenerateMBR] No unit operations parsed. Aborting."
        GoTo CleanExit
    End If

    Dim templateTitle As String
    Dim awarenessSuffix As String
    templateTitle = "Laborvorschrift Template"
    awarenessSuffix = "weight_name"


    Debug.Print "[GenerateMBR] Using template: " & templateTitle
    Call BuildMBRFromUnitOperations(unitOperations, templateTitle, awarenessSuffix)
    Debug.Print "[GenerateMBR] MBR generation finished."

CleanExit:
    Debug.Print "[Timing] GenerateMBR Duration: " & format(Timer - tStart, "0.000") & " sec"
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Debug.Print "[GenerateMBR] Error: " & Err.Description
    MsgBox "An error occurred while generating the MBR: " & Err.Description, vbCritical, "Flow Chart Error"
    Err.Clear
    GoTo CleanExit
End Sub


' === Modified: BuildFlowChartFromUnitOperations ===
Public Sub BuildMBRFromUnitOperations(ByVal unitOps As Collection, ByVal templateTitle As String, ByVal awarenessSuffix As String)
    On Error GoTo ErrorHandler
    Debug.Print "[GenerateMBR] Starting flow chart generation..."

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
        Debug.Print "[GenerateMBR] Error: Template zone not found: " & templateTitle
        Exit Sub
    End If

    Set ccFlowChart = FindContentControlByTitle("Laborvorschrift") 'For generalization, pass as argument
    If ccFlowChart Is Nothing Then
        MsgBox "Error: 'Laborvorschrift' content control not found.", vbCritical, "Flow Chart Error"
        Debug.Print "[GenerateMBR] Error: Output zone not found."
        Exit Sub
    End If

    On Error Resume Next
    ccFlowChart.Range.Delete
    If Err.number <> 0 Then
        Debug.Print "[GenerateMBR] Warning: Unable to clear Laborvorschrift range - " & Err.Description
        Err.Clear
    Else
        Debug.Print "[GenerateMBR] Cleared existing content in 'Laborvorschrift'."
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
                    Call PopulateMBRBlock(unitOp, awarenessSuffix)
                Else
                    Debug.Print "[GenerateMBR] Warning: Inserted block not found for: " & unitOp.title
                End If

                foundTemplate = True
                Exit For
            End If
        Next ccTemplateBlock

        If Not foundTemplate Then
            MsgBox "Template block for '" & unitOp.title & "' not found.", vbExclamation, "Missing Template"
            Debug.Print "[GenerateMBR] Warning: Template block not found for unit operation titled '" & unitOp.title & "'."
        End If
    Next i

    Debug.Print "[GenerateMBR] Flow chart generation completed."
    Call FinalizeMBRLayout ' to generalize, specific finalization, based on flow chart or MBR
    Exit Sub

ErrorHandler:
    Debug.Print "[GenerateMBR] Error: " & Err.Description
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

Private Sub PopulateMBRBlock(unitOp As clsUnitOperation, Optional defaultAwareness As String = "flowchart")
    On Error GoTo ErrorHandler

    Dim flowChartCC As ContentControl
    Dim nestedCC As ContentControl
    Dim fullTag As String
    Dim valueText As String
    Dim wasFound As Boolean
    Dim foundBlock As Boolean

    If unitOp Is Nothing Then
        Debug.Print "[GenerateMBR] Error: unitOp is Nothing"
        Exit Sub
    End If
    If Trim(unitOp.id) = "" Then
        Debug.Print "[GenerateMBR] Error: unitOp.ID is empty"
        Exit Sub
    End If

    foundBlock = False

    ' We learned that this does not work with any content control outside richt text content control
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
        Debug.Print "[GenerateMBR] Error: Could not locate block with Tag = " & unitOp.id & " to populate."
    End If

    Exit Sub

ErrorHandler:
    Debug.Print "[GenerateMBR] ERROR #" & Err.number & " at line " & Erl & ": " & Err.Description
    Err.Clear
End Sub

Public Sub FinalizeMBRLayout()
    On Error GoTo ErrorHandler
    Debug.Print "[GenerateMBR] Starting MBR finalization..."

    Dim tStart As Double
    tStart = Timer
    Application.ScreenUpdating = False ' ? Boost performance

    Dim ccFlowChart As ContentControl
    Dim ccBlock As ContentControl
    Dim tbl As Table
    Dim para As Paragraph
    Dim firstParaStart As Long, lastParaEnd As Long

    Set ccFlowChart = FindContentControlByTitle("Laborvorschrift") 'For generalization pass as argument
    If ccFlowChart Is Nothing Then
        Debug.Print "[GenerateMBR] Error: 'Laborvorschrift' content control not found."
        GoTo CleanExit
    End If

    firstParaStart = ccFlowChart.Range.Paragraphs(1).Range.Start
    lastParaEnd = ccFlowChart.Range.Paragraphs(ccFlowChart.Range.Paragraphs.count).Range.End

    ' NOTE: At some point the CC get flattened anyways, this is an unwanted feature/bug
    'Debug.Print "[GenerateMBR] Call Flattened all middle-layer controls in Flow Chart."
   ' Call FlattenMiddleLayerContentControls(ccFlowChart)
    'Debug.Print "[GenerateMBR] Flattened all middle-layer controls in Flow Chart."
    Debug.Print "[GenerateMBR][SKIPPED] Flattened all middle-layer controls in Flow Chart."

    For Each para In ccFlowChart.Range.Paragraphs
        If para.Range.Tables.count = 0 And Len(Trim(para.Range.text)) > 0 Then
            If para.Range.Start > firstParaStart And para.Range.End < lastParaEnd Then
                para.Range.Delete
                ' Debug.Print "[GenerateMBR] Removed line break or non-table paragraph."
            End If
        End If
    Next para

    Debug.Print "[GenerateMBR] MBR finalization completed."

CleanExit:
Debug.Print "[Timing] FinalizeFlowChartLayout Duration: " & format(Timer - tStart, "0.000") & " sec"
    Application.ScreenUpdating = True ' ? Restore state
    Exit Sub

ErrorHandler:
    Debug.Print "[GenerateMBR] Error in FinalizeFlowChartLayout: " & Err.Description
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

Public Sub RefreshMBR()
    On Error GoTo FatalError

    Const MBR_TITLE As String = "Laborvorschrift"  ' <-- adjust if your CC title is different

    Dim t0 As Double, tParse As Double, tWalk As Double
    t0 = Timer

    Dim doc As Document: Set doc = ThisDocument
    Dim prevSU As Boolean: prevSU = Application.ScreenUpdating
    Dim prevTrack As Boolean: prevTrack = doc.TrackRevisions
    Dim prevAlerts As WdAlertLevel: prevAlerts = Application.DisplayAlerts

    Debug.Print String(60, "-")
    Debug.Print "[RefreshMBR] Begin refresh (" & MBR_TITLE & ")."

    ' Guards
    Application.ScreenUpdating = False
    Application.DisplayAlerts = wdAlertsNone
    DisableEvents
    doc.TrackRevisions = False

    ' Locate MBR container
    Dim ccMBR As ContentControl
    Set ccMBR = FindContentControlByTitle(MBR_TITLE)
    If ccMBR Is Nothing Then
        Debug.Print "[RefreshMBR] ERROR: '" & MBR_TITLE & "' content control not found."
        MsgBox "'" & MBR_TITLE & "' area not found.", vbExclamation, "Refresh MBR"
        GoTo Cleanup
    End If

    ' Parse source
    Dim unitOps As Collection
    Set unitOps = ParseProcessDescription()
    If unitOps Is Nothing Or unitOps.count = 0 Then
        Debug.Print "[RefreshMBR] WARNING: ParseProcessDescription returned no unit operations."
        MsgBox "No unit operations parsed. Nothing to refresh.", vbExclamation, "Refresh MBR"
        GoTo Cleanup
    End If

    tParse = Timer
    Debug.Print "[RefreshMBR] Parsed unit operations in " & format(tParse - t0, "0.000") & " s"
    Debug.Print "[RefreshMBR] UnitOps count = " & unitOps.count

    ' Walk direct children CCs of the MBR container
    Dim total As Long, updated As Long, naNotFound As Long
    Dim naMalformed As Long, naMissingUO As Long, perItemErrs As Long

    Dim cc As ContentControl
    For Each cc In ccMBR.Range.contentControls
        total = total + 1
        On Error GoTo ItemError

        Dim tg As String: tg = Trim$(cc.tag)
        If Len(tg) < 7 Then
            ' Need at least 7 chars (5 for unit op id, + 2 for field id)
            SafeSetCCText cc, "N/A"
            naMalformed = naMalformed + 1
            GoTo NextCC
        End If

        ' First 5 chars: unit operation id
        Dim unitOpId As String
        unitOpId = Left$(tg, 5)

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
            SafeSetCCText cc, valueText
            updated = updated + 1
        Else
            SafeSetCCText cc, "N/A"
            naNotFound = naNotFound + 1
        End If

        GoTo NextCC

ItemError:
        perItemErrs = perItemErrs + 1
        Debug.Print "[RefreshMBR][ItemError] Title='" & cc.title & "' Tag='" & cc.tag & _
                    "' :: " & Err.number & " - " & Err.Description
        Err.Clear
        On Error GoTo FatalError
        SafeSetCCText cc, "N/A"

NextCC:
        ' continue
    Next cc

    tWalk = Timer
    Debug.Print "[RefreshMBR] Walked " & total & " CC(s) in " & format(tWalk - tParse, "0.000") & " s"
    Debug.Print "  Updated:                    " & updated
    Debug.Print "  N/A (not found in model):   " & naNotFound
    Debug.Print "  N/A (malformed tag):        " & naMalformed
    Debug.Print "  N/A (missing unit op):      " & naMissingUO
    Debug.Print "  Item errors:                " & perItemErrs
    Debug.Print "[RefreshMBR] TOTAL: " & format(Timer - t0, "0.000") & " s"
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
    Debug.Print "[RefreshMBR][FATAL] " & Err.number & " - " & Err.Description
    MsgBox "Error during MBR refresh: " & Err.Description, vbCritical, "Refresh MBR"
    Resume Cleanup
End Sub





