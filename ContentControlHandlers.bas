Attribute VB_Name = "ContentControlHandlers"
Option Explicit
' In a shared module, e.g., ContentControlHandlers
Public Function ExpandContentControl(ByVal cc As ContentControl, ByRef doc As Document) As Boolean
    Dim didProcess As Boolean: didProcess = False
    Dim tSection As Single

    Do
        If Not ContentControlHelpers.ParentContentControlTitled("ProcessDescription", cc) Then Exit Do

        Select Case LCase(cc.title)
            Case "input", "product"
                tSection = Timer
                HandleCompoundInputEnter cc
                didProcess = True

            Case "waste"
                tSection = Timer
                HandleWasteInputEnter cc
                didProcess = True
        End Select
    Loop While False

    If didProcess Then doc.UndoClear
    ExpandContentControl = didProcess
End Function


Public Function CollapseContentControl(ByVal cc As ContentControl, ByRef doc As Document) As Boolean
    Dim didProcess As Boolean: didProcess = False
    Dim tSection As Single

    If Not ContentControlHelpers.ParentContentControlTitled("ProcessDescription", cc) Then Exit Function

    Select Case LCase(cc.title)
        Case "input", "product"
            tSection = Timer
            HandleCompoundInputExit cc
            Debug.Print "Time - HandleCompoundInputExit: " & format(Timer - tSection, "0.000") & "s"

            If IsTableUpdateEnabled() Then
                tSection = Timer: PopulateBOMTable
                Debug.Print "Time - PopulateBOMTable: " & format(Timer - tSection, "0.000") & "s"
            End If

            If IsTableUpdateEnabled() Then
                tSection = Timer: PopulateMassBalanceTable
                Debug.Print "Time - PopulateMassBalanceTable: " & format(Timer - tSection, "0.000") & "s"
            End If
            didProcess = True

        Case "waste"
            tSection = Timer
            HandleWasteInputExit cc
            Debug.Print "Time - HandleWasteInputExit: " & format(Timer - tSection, "0.000") & "s"

            If IsTableUpdateEnabled() Then
                tSection = Timer: PopulateWSTable
                Debug.Print "Time - PopulateWSTable: " & format(Timer - tSection, "0.000") & "s"
            End If

            If IsTableUpdateEnabled() Then
                tSection = Timer: PopulateMassBalanceTable
                Debug.Print "Time - PopulateMassBalanceTable: " & format(Timer - tSection, "0.000") & "s"
            End If
            didProcess = True
    End Select

    CollapseContentControl = didProcess
End Function

' Computes model update from user edit string, no UI mutation
Public Sub ComputeCompoundEdit(cc As ContentControl)
    On Error GoTo HandleError

    Dim cmp As Compound
    Dim editStr As String
    Dim ref As Compound

    Set cmp = GetGlobalCompoundCollection().GetById(cc.tag)
    If cmp Is Nothing Then
        Debug.Print "[ERROR] No compound found for tag: " & cc.tag
        Exit Sub
    End If

    editStr = cc.Range.text
    Set ref = GetGlobalCompoundCollection().GetReferenceCompound()

    Call CompoundMutator.ApplyEditString(cmp, editStr, ref)
    Exit Sub

HandleError:
    Debug.Print "[ERR] ComputeCompoundEdit: " & Err.Description
    Err.Clear
End Sub

' Updates the UI (display string + bold formatting), assumes model is already up to date
Public Sub UpdateCompoundContentControlDisplay(cc As ContentControl)
    On Error GoTo HandleError

    Dim cmp As Compound
    Dim ref As Compound

    Set cmp = GetGlobalCompoundCollection().GetById(cc.tag)
    Set ref = GetGlobalCompoundCollection().GetReferenceCompound()

    Application.ScreenUpdating = False
    cc.Range.text = cmp.ToDisplayString("low", ref)
    cc.Range.Font.Bold = True
    Application.ScreenUpdating = True
    Exit Sub

HandleError:
    Debug.Print "[ERR] UpdateCompoundContentControlDisplay: " & Err.Description
    Err.Clear
End Sub

' MOVE to
' Safely set the visible text of a content control without crashing on locks/protection.
' - Only writes to Text or RichText CCs.
' - Temporarily clears LockContents, then restores it.
' - Falls back to writing directly into the range for plain text.
Public Sub SafeSetCCText(ByVal cc As ContentControl, ByVal valueText As String)
    On Error GoTo Fail

    ' Only text-like CCs should be written via .Range.Text
    Select Case cc.Type
        Case wdContentControlText, wdContentControlRichText
            Dim wasLocked As Boolean
            wasLocked = cc.LockContents
            If wasLocked Then cc.LockContents = False

            ' Standard write
            cc.Range.text = valueText

            If wasLocked Then cc.LockContents = True

        Case Else
            ' Non-text CC (picture, building block, etc.) ? do not attempt to write
            ' Your policy is to label N/A on invalid fields; show it next to CC if needed:
            ' Optional: cc.Range.Text = "" : Not recommended for non-text types.
            ' Here we just skip silently.
    End Select
    Exit Sub

Fail:
    ' Last-resort: avoid raising during bulk refresh. Do not touch structure.
    Debug.Print "[SafeSetCCText][WARN] Could not set text for CC Title='" & cc.title & _
                "', Tag='" & cc.tag & "': " & Err.number & " - " & Err.Description
    Err.Clear
End Sub


