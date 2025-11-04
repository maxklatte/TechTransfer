Attribute VB_Name = "CompoundHandler"
' Module: CompoundHandler
Option Explicit

Public gStoffdatenbank As Stoffdatenbank
Public gCompounds As CompoundCollection

Public Sub HandleCompoundInputEnter(cc As ContentControl)
    On Error GoTo HandleError

    ' Validate tag
    Dim id As String: id = cc.tag
    If Len(id) = 0 Then
        Debug.Print "No tag set on ContentControl."
        Exit Sub
    End If

    ' Try to retrieve existing Compound
    Dim cmp As Compound
    Set cmp = GetGlobalCompoundCollection().GetById(id)
    Debug.Print "Compound lookup by ID: " & id & IIf(cmp Is Nothing, " not found.", " found.")

    If cmp Is Nothing Then
        Debug.Print "Creating Compound for ID: " & id
        Dim temp As Compound
        Set temp = New Compound

        Select Case LCase(cc.title)
            Case "input"
                Set cmp = temp.BuildDefaultWaterCompound(id)
                Debug.Print "Built default water compound."

            Case "product"
                Set cmp = temp.BuildProduct(id)
                Debug.Print "Built product compound."
        End Select

        GetGlobalCompoundCollection().Add cmp
        Debug.Print "Compound added to global collection."
    Else
        Debug.Print "Compound already exists for ID: " & id
    End If

    ' Generate display text
    Dim displayText As String
    On Error GoTo DisplayError
    displayText = cmp.ToDisplayString("edit", GetGlobalCompoundCollection().GetReferenceCompound())
    On Error GoTo HandleError

    ' Update ContentControl text
    Application.ScreenUpdating = False
    cc.Range.text = displayText
    Application.ScreenUpdating = True

    Exit Sub

DisplayError:
    Debug.Print "[ERR] ToDisplayString failed for ID: " & id & " - " & Err.Description
    displayText = "[ERROR: ToDisplayString failed]"
    Err.Clear
    Resume Next

HandleError:
    Debug.Print "HandleCompoundInputEnter error: " & Err.Description
    Err.Clear
End Sub




' Purpose: Simplified fully to use CompoundMutator smart pipeline, with user conflict notification

' Legacy exit handler, composes compute + display
Public Sub HandleCompoundInputExit(cc As ContentControl)
    On Error GoTo HandleError

    Call ComputeCompoundEdit(cc)
    Call UpdateCompoundContentControlDisplay(cc)
    Exit Sub

HandleError:
    Debug.Print "[ERR] HandleCompoundInputExit: " & Err.Description
    Err.Clear
End Sub

Public Function GetGlobalStoffdatenbank() As Stoffdatenbank
    If gStoffdatenbank Is Nothing Then
        Set gStoffdatenbank = New Stoffdatenbank
        If Not gStoffdatenbank.LoadIfNotInitialized Then
            Set gStoffdatenbank = Nothing
        End If
    End If
    Set GetGlobalStoffdatenbank = gStoffdatenbank
End Function



Public Function GetGlobalCompoundCollection() As CompoundCollection
    If gCompounds Is Nothing Then
        Set gCompounds = New CompoundCollection
        If Not gCompounds.LoadFromWord Then
            Debug.Print "[INIT] CompoundCollection: No existing XML found."
        End If
        If Not gCompounds.Validate Then
            Debug.Print "[WARN] CompoundCollection failed validation after load."
        End If
    End If
    Set GetGlobalCompoundCollection = gCompounds
End Function

Public Sub RefreshCompoundDisplays()
    Dim i As Long
    Dim cc As ContentControl
    Dim cmp As Compound

    Dim oldState As Boolean: oldState = Application.ScreenUpdating
    Application.ScreenUpdating = False

    With GetGlobalCompoundCollection()
        For i = 1 To .count
            Set cmp = .item(i)
            Set cc = ContentControlHelpers.FindControlByTag(cmp.id)
            If Not cc Is Nothing Then
                cc.Range.text = cmp.ToDisplayString("low", GetGlobalCompoundCollection().GetReferenceCompound())
            Else
                Debug.Print "[WARN] No ContentControl found for compound ID: " & cmp.id
            End If
        Next i
    End With

    Application.ScreenUpdating = oldState
End Sub
