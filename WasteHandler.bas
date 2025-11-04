Attribute VB_Name = "WasteHandler"
'=== MODULE: WasteHandler ===
Option Explicit

Public gWasteCollection As WasteCollection

Public Sub HandleWasteInputEnter(cc As ContentControl)
    On Error GoTo HandleError

    Dim id As String: id = cc.tag
    If Len(id) = 0 Then
        Debug.Print "[WARN] No tag set on ContentControl."
        Exit Sub
    End If

    Dim w As Waste
    Set w = GetGlobalWasteCollection().GetById(id)

    If w Is Nothing Then
        Debug.Print "[INFO] Waste not found for ID: " & id
        Set w = BuildWaste(id)
        GetGlobalWasteCollection().Add w
    End If

    Call DisplayWasteInEditMode(cc)
    Exit Sub

HandleError:
    Debug.Print "[ERROR] HandleWasteInputEnter: " & Err.Description
    Err.Clear
End Sub


Public Sub HandleWasteInputExit(cc As ContentControl)
    On Error GoTo HandleError

    Call ComputeWasteEdit(cc)
    Call UpdateWasteContentControlDisplayLow(cc)
    Exit Sub

HandleError:
    Debug.Print "[ERROR] HandleWasteInputExit: " & Err.Description
    Err.Clear
End Sub




Public Function GetGlobalWasteCollection() As WasteCollection
    If gWasteCollection Is Nothing Then
        Set gWasteCollection = New WasteCollection
        If Not gWasteCollection.LoadFromWord Then
            Debug.Print "[INIT] WasteCollection: No existing XML found."
        End If
        If Not gWasteCollection.Validate Then
            Debug.Print "[WARN] WasteCollection failed validation after load."
        End If
    End If
    Set GetGlobalWasteCollection = gWasteCollection
End Function

Private Function BuildWaste(tagId As String) As Waste
    Dim w As Waste
    Set w = New Waste
    Set BuildWaste = w.BuildWaste(tagId)
End Function

Public Sub RefreshWasteDisplays()
    Dim i As Long
    Dim cc As ContentControl
    Dim w As Waste

    Dim oldState As Boolean: oldState = Application.ScreenUpdating
    Application.ScreenUpdating = False

    With GetGlobalWasteCollection()
        For i = 1 To .count
            Set w = .item(i)
            Set cc = ContentControlHelpers.FindControlByTag(w.id)
            If Not cc Is Nothing Then
                cc.Range.text = w.ToDisplayString("low")
            Else
                Debug.Print "[WARN] No ContentControl found for waste ID: " & w.id
            End If
        Next i
    End With

    Application.ScreenUpdating = oldState
End Sub

Public Sub ComputeWasteEdit(cc As ContentControl)
    On Error GoTo HandleError

    Dim w As Waste
    Dim editStr As String

    Set w = GetGlobalWasteCollection().GetById(cc.tag)
    If w Is Nothing Then
        MsgBox "No waste found for tag: " & cc.tag, vbCritical
        Exit Sub
    End If

    editStr = cc.Range.text
    Call w.ApplyEditString(editStr)
    Exit Sub

HandleError:
    MsgBox "Error in ComputeWasteEdit: " & Err.Description, vbCritical
    Err.Clear
End Sub
Public Sub UpdateWasteContentControlDisplayLow(cc As ContentControl)
    On Error GoTo HandleError

    Dim w As Waste
    Set w = GetGlobalWasteCollection().GetById(cc.tag)
    If w Is Nothing Then Exit Sub

    Application.ScreenUpdating = False
    cc.Range.text = w.ToDisplayString("low")
    cc.Range.Font.Bold = True
    Application.ScreenUpdating = True
    Exit Sub

HandleError:
    MsgBox "Error displaying waste in 'low' format: " & Err.Description, vbCritical
    Err.Clear
End Sub
Public Sub DisplayWasteInEditMode(cc As ContentControl)
    On Error GoTo HandleError

    Dim w As Waste
    Set w = GetGlobalWasteCollection().GetById(cc.tag)
    If w Is Nothing Then Exit Sub

    Application.ScreenUpdating = False
    cc.Range.text = w.ToDisplayString("edit")
    Application.ScreenUpdating = True
    Exit Sub

HandleError:
    MsgBox "Error displaying waste in 'edit' format: " & Err.Description, vbCritical
    Err.Clear
End Sub




