VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsUnitOperation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' === Class Module: clsUnitOperation ===
Option Explicit

' Core properties
Public Step As String ' Formerly ID
Public id As String    ' Extracted from Tag (first 5 characters)
Public tag As String   ' ContentControl.Tag
Public title As String ' ContentControl.Title
Public InstructionText As String

' Component Collections
Public Inputs As Collection
Public Outputs As Collection
Public Equipment As Collection
Public Parameters As Collection

' Transient Row object reference to source table row
' Can be ditched
Public sourceRow As row

    ' New: Track original source ID (for rendering clone linkage)
    Public SourceID As String

Private Sub Class_Initialize()
    Set Inputs = New Collection
    Set Outputs = New Collection
    Set Equipment = New Collection
    Set Parameters = New Collection
End Sub

Public Sub AddInput(text As String, tag As String)
    Dim inputObj As Object
    Set inputObj = CreateObject("Scripting.Dictionary")
    inputObj.Add "Text", text
    inputObj.Add "Tag", tag
    Inputs.Add inputObj
End Sub

Public Sub AddOutput(text As String, tag As String)
    Dim outputObj As Object
    Set outputObj = CreateObject("Scripting.Dictionary")
    outputObj.Add "Text", text
    outputObj.Add "Tag", tag
    Outputs.Add outputObj
End Sub

Public Sub AddEquipment(text As String, tag As String)
    Dim equipObj As Object
    Set equipObj = CreateObject("Scripting.Dictionary")
    equipObj.Add "Text", text
    equipObj.Add "Tag", tag
    Equipment.Add equipObj
End Sub

Public Sub AddParameter(name As String, value As Variant, tag As String)
    Dim param As Object
    Set param = CreateObject("Scripting.Dictionary")
    param.Add "Name", name
    param.Add "Value", value
    param.Add "Tag", tag
    Parameters.Add param
End Sub

' === Class Module: clsUnitOperation ===

Public Function Clone(newID As String) As clsUnitOperation
    Dim clonedOp As clsUnitOperation
    Set clonedOp = New clsUnitOperation

    ' === Scalar Fields ===
    clonedOp.Step = Me.Step ' Step will be renumbered later
    clonedOp.id = newID
    clonedOp.tag = newID & Right(Me.tag, 2)
    clonedOp.title = Me.title
    clonedOp.InstructionText = Me.InstructionText
    clonedOp.SourceID = Me.id ' Track where it came from

    ' === Deep or Reference Clone Based on Type ===
    Call CloneCollectionWithNewTags(Me.Inputs, clonedOp.Inputs, newID, "inputs", True)
    Call CloneCollectionWithNewTags(Me.Outputs, clonedOp.Outputs, newID, "outputs", True)
    Call CloneCollectionWithNewTags(Me.Equipment, clonedOp.Equipment, newID, "equipment", False)
    Call CloneCollectionWithNewTags(Me.Parameters, clonedOp.Parameters, newID, "parameters", False)

    ' === Maintain source row for visual rendering ===
    ' Can be ditched...
    Set clonedOp.sourceRow = Me.sourceRow


    Set Clone = clonedOp
End Function


Public Sub PrintSummary()
    Dim item As Variant
    Debug.Print "-------------------------------------------------------------"
    Debug.Print "Unit Operation Step: " & Step
    Debug.Print "ID (Real): " & id
    Debug.Print "Title: " & title
    Debug.Print "Tag: " & tag
'    Debug.Print "Instruction: " & InstructionText

    Debug.Print "Inputs:"
    For Each item In Inputs
        Debug.Print item("Tag") & ": " & item("Text")
        'Debug.Print "  - " & item("Text") & " [Tag: " & item("Tag") & "]"
    Next item

    Debug.Print "Outputs:"
    For Each item In Outputs
            Debug.Print item("Tag") & ": " & item("Text")
        'Debug.Print "  - " & item("Text") & " [Tag: " & item("Tag") & "]"
    Next item

    Debug.Print "Equipment:"
    For Each item In Equipment
            Debug.Print item("Tag") & ": " & item("Text")
        'Debug.Print "  - " & item("Text") & " [Tag: " & item("Tag") & "]"
    Next item

    Debug.Print "Parameters:"
    For Each item In Parameters
            Debug.Print item("Tag") & ": " & item("Value")
        'Debug.Print "  - " & item("Name") & ": " & item("Value") & " [Tag: " & item("Tag") & "]"
    Next item
End Sub

Public Function GetTextByTag(fullTag As String, ByRef found As Boolean) As String
    Dim item As Object
    found = False
    GetTextByTag = "N/A"
    
    Select Case Right(fullTag, 2)
        Case "00"
            found = True
            GetTextByTag = Me.title
            Exit Function
        Case "01"
            found = True
            GetTextByTag = Me.Step
            Exit Function
        Case "99"
            found = True
            GetTextByTag = Me.InstructionText
            Exit Function
    End Select

    ' Compound-aware fallback
    Dim cmpText As String
    cmpText = GetGlobalCompoundCollection().GetTextByTag(fullTag, found)
    If found Then
        GetTextByTag = cmpText
        Exit Function
    End If

    ' Legacy item search
    For Each item In Inputs
        If item("Tag") = fullTag Then
            found = True
            GetTextByTag = item("Text")
            Exit Function
        End If
    Next item

    For Each item In Outputs
        If item("Tag") = fullTag Then
            found = True
            GetTextByTag = item("Text")
            Exit Function
        End If
    Next item

    For Each item In Equipment
        If item("Tag") = fullTag Then
            found = True
            GetTextByTag = item("Text")
            Exit Function
        End If
    Next item

    For Each item In Parameters
        If item("Tag") = fullTag Then
            found = True
            GetTextByTag = item("Value")
            Exit Function
        End If
    Next item
End Function

Public Function ToString() As String
    ToString = "[" & id & "] " & title
End Function

Public Function IsValid() As Boolean
    Dim isOk As Boolean: isOk = True

    If Len(Trim(Me.id)) <> 5 Then isOk = False
    If Len(Trim(Me.title)) = 0 Then isOk = False
    If Len(Trim(Me.InstructionText)) = 0 Then isOk = False

    Dim item As Variant
    For Each item In Me.Inputs
        If Not item.Exists("Text") Or Len(Trim(item("Text"))) = 0 Then isOk = False
    Next

    IsValid = isOk
End Function
