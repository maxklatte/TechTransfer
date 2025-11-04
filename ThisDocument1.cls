VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' ThisDocument Module
Option Explicit

' In ThisDocument
Dim appEvents As clsAppEvents

Private Sub Document_Open()
    On Error GoTo ErrorHandler

    Set appEvents = New clsAppEvents
    If Not appEvents Is Nothing Then
        Set appEvents.appWord = word.Application
    End If

    Exit Sub
ErrorHandler:
    MsgBox "Startup error: " & Err.Description, vbExclamation
    Resume Next
End Sub



Private Sub Document_ContentControlOnExit(ByVal cc As ContentControl, Cancel As Boolean)
    On Error GoTo ErrorHandler

    If Not IsEventsEnabled() Then Exit Sub
    If Not IsUserInteractionEnabled() Then Exit Sub

    Dim doc As Document
    Dim wasTracked As Boolean
    Dim didProcess As Boolean: didProcess = False
    Dim tStart As Single, tSection As Single

    tStart = Timer
    Application.ScreenUpdating = False

    Set doc = cc.Range.Document
    wasTracked = doc.TrackRevisions
    doc.TrackRevisions = False

    If cc.title = "UnitOperationDropdown" Then
        tSection = Timer
        UnitOperationManager.HandleDropdownSelection cc
        Debug.Print "Time - UnitOperationDropdown: " & format(Timer - tSection, "0.000") & "s"
        didProcess = True
    End If

    If CollapseContentControl(cc, doc) Then didProcess = True
    If didProcess Then doc.UndoClear

Cleanup:
    On Error Resume Next                 ' be defensive in cleanup
    If Not doc Is Nothing Then
        If Selection.StoryType = wdMainTextStory Then
            doc.TrackRevisions = wasTracked
        End If
    End If
    Application.ScreenUpdating = True
    Debug.Print "Total Time - ContentControlOnExit [" & cc.title & "]: " & format(Timer - tStart, "0.000") & "s"
    Exit Sub

ErrorHandler:
    Debug.Print "!! Error in OnExit [" & cc.title & "]: " & Err.number & " - " & Err.Description
    MsgBox "An error occurred while processing '" & cc.title & "'." & vbCrLf & _
           "Details: " & Err.Description, vbExclamation, "Macro Error"
    Resume Cleanup
End Sub



Private Sub Document_ContentControlOnEnter(ByVal cc As ContentControl)
    Dim doc As Document
    Dim wasTracked As Boolean
    Dim tStart As Single

    On Error GoTo ErrorHandler
    
    If Not IsEventsEnabled() Then Exit Sub
    If Not IsUserInteractionEnabled() Then Exit Sub

    Set doc = cc.Range.Document
    wasTracked = doc.TrackRevisions
    doc.TrackRevisions = False
    tStart = Timer

    Call ExpandContentControl(cc, doc)

Cleanup:
    doc.TrackRevisions = wasTracked
    Debug.Print "Total Time - ContentControlOnEnter [" & cc.title & "]: " & format(Timer - tStart, "0.000") & "s"
    Exit Sub

ErrorHandler:
    Debug.Print "!! Error in OnEnter [" & cc.title & "]: " & Err.number & " - " & Err.Description
    MsgBox "An error occurred while entering '" & cc.title & "'." & vbCrLf & _
           "Details: " & Err.Description, vbExclamation, "Macro Error"
    Resume Cleanup
End Sub





Private Sub Document_ContentControlBeforeDelete(ByVal cc As ContentControl, ByVal InUndoRedo As Boolean)
    On Error GoTo ErrorHandler

    'If Not IsUserInteractionEnabled() Then Exit Sub
    If Not IsEventsEnabled() Then Exit Sub

    If DeletedElementIsUnitOperationID(cc.tag) And ParentContentControlTitled("ProcessDescription", cc) Then
        Debug.Print "Unit operation deleted."

        ' Defer post-delete logic
        Application.OnTime Now + timeValue("00:00:01"), "ContentControlAfterDelete_ContentControlHelpers"
    End If

    Exit Sub

ErrorHandler:
    Debug.Print "!! Error in BeforeDelete [" & cc.title & "]: " & Err.number & " - " & Err.Description
    Resume Next
End Sub



Private Function DeletedElementIsUnitOperationID(ByVal tag As String) As Boolean
    On Error GoTo ErrorHandler
    'Debug.Print "Validating tag: " & tag
    
    ' Ensure the tag is a valid string
    If Len(tag) = 7 And Right(tag, 2) = "01" Then
        DeletedElementIsUnitOperationID = True
        ' Debug.Print "Tag is a valid Unit Operation ID: " & tag
    Else
        DeletedElementIsUnitOperationID = False
        'Debug.Print "Tag is not a valid Unit Operation ID: " & tag
    End If
    Exit Function

ErrorHandler:
    Debug.Print "Error in DeletedElementIsUnitOperationID: " & Err.Description
    Err.Clear
    DeletedElementIsUnitOperationID = False
End Function


' Ribbon Callback Handlers for ThisDocument
' === SHARED SAFETY WRAPPER ===
Public Sub SafeCall(control As IRibbonControl)
    Dim actionLabel As String
    actionLabel = control.tag
    
    Dim doc As Document: Set doc = ThisDocument
    Dim wasTracked As Boolean: wasTracked = doc.TrackRevisions
    Dim didProcess As Boolean: didProcess = False

    On Error GoTo ErrorHandler
    DisableEvents
    Application.ScreenUpdating = False
    doc.TrackRevisions = False

    Debug.Print "[Ribbon] " & actionLabel

Select Case actionLabel

    ' === PROCESS BUILDER ===
    Case "AddUnitOperation":              UserMacros.AddUnitOperation
    Case "InsertUnitOperationDropDown":   UserMacros.AddDropdownToTableRow
    Case "CloneUnitOps":                  UserMacros.CloneUnitOpsDialog

    ' === CHEMISTRY ===
    Case "SelectReferenceCompound":       UserMacros.SelectReferenceCompound
    Case "ScaleUp":                       UserMacros.ScaleUp

    ' === PROCESS VISUALIZATION ===
    Case "ForceUpdateTables"
        Renumber
        PruneAllCollections
        PopulateWSTable
        PopulateBOMTable
        PopulateBOMSummaryTable
        PopulateMassBalanceTable
        PopulateIPCsTable
        PopulateHoldPointsTable
        PopulateClearPointsTable

    ' === FIELD OPERATIONS ===
    Case "Expand":                        UserMacros.ExpandField
    Case "Compute":                       UserMacros.ComputeField
    Case "Collapse":                      UserMacros.CollapseField
    Case "Refresh":                       UserMacros.RefreshStoffdatenForField

    ' === PROCESS TRANSLATOR ===
    Case "GenerateFlowChart":             FlowChartCreator.GenerateFlowChart
    Case "GenerateLaborvorschrift":       MBRCreator.GenerateMBR

    ' === MACRO CONTROL ===
    Case "ForceReloadStoffdatenbank":     UserMacros.ForceReloadStoffdatenbank
    Case "ToggleTableUpdates":            UserMacros.ToggleTableUpdateState
    Case "ToggleUserInteraction":         UserMacros.ToggleUserInteraction

    ' === FALLBACK ===
    Case Else
        Err.Raise vbObjectError + 911, , "Unknown action label: " & actionLabel
End Select


    didProcess = True
    If didProcess Then doc.UndoClear

Cleanup:
    doc.TrackRevisions = wasTracked
    Application.ScreenUpdating = True
    EnableEvents
    Exit Sub

ErrorHandler:
    MsgBox "[ERROR] " & actionLabel & vbCrLf & _
           """" & Err.Description & """", vbCritical, "Unexpected Error"
    Resume Cleanup
End Sub

