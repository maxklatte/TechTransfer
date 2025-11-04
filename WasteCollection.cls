VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WasteCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=== CLASS MODULE: WasteCollection ===
Option Explicit

Private wastes As Collection

Private Sub Class_Initialize()
    Set wastes = New Collection
End Sub

Public Sub Add(ByVal w As Waste)
    wastes.Add w, w.id
    Me.SortByStepNumber
End Sub

Public Sub RemoveById(ByVal wasteId As String)
    On Error Resume Next
    wastes.Remove wasteId
End Sub

Public Function GetById(ByVal wasteId As String) As Waste
    On Error Resume Next
    Set GetById = wastes(wasteId)
    If GetById Is Nothing Then
        Debug.Print "[WARN] Waste not found: " & wasteId
    End If
End Function

Public Property Get count() As Long
    count = wastes.count
End Property

Public Property Get item(index As Variant) As Waste
    Set item = wastes(index)
End Property

Public Function ToXml() As String
    Dim xmlDoc As Object: Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    Dim root As Object: Set root = xmlDoc.createElement("Wastes")
    root.setAttribute "xmlns", "urn:my-waste-namespace"
    xmlDoc.appendChild root

    Dim w As Waste, node As Object
    For Each w In wastes
        Set node = w.ToXmlNode(xmlDoc)
        root.appendChild node
    Next w

    ToXml = xmlDoc.xml
End Function

Public Sub LoadFromXml(xmlString As String)
    Dim xmlDoc As Object: Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    xmlDoc.LoadXML xmlString

    Dim nodes As Object, node As Object
    Set nodes = xmlDoc.SelectNodes("//Waste")

    Dim w As Waste
    For Each node In nodes
        Set w = New Waste
        w.FromXmlNode node
        Add w
    Next node
        Me.SortByStepNumber
End Sub

Public Sub SaveToWord()
    On Error GoTo ErrHandler

    Dim doc As Document: Set doc = ThisDocument
    Dim part As Office.CustomXMLPart
    Dim xmlString As String
    Dim ns As String: ns = "urn:my-waste-namespace"

    For Each part In doc.CustomXMLParts
        If InStr(1, part.xml, ns, vbTextCompare) > 0 Then
            part.Delete
            Exit For
        End If
    Next part

    xmlString = Me.ToXml

    If Not IsWellFormedXml(xmlString) Then
        Debug.Print xmlString
        Err.Raise vbObjectError + 1002, , "Generated Waste XML is not well-formed."
    End If

    doc.CustomXMLParts.Add xmlString
    'MsgBox "Waste XML saved successfully.", vbInformation
    Exit Sub

ErrHandler:
    Debug.Print "[SaveToWord Error - Waste]"
    Debug.Print "Line: " & Erl
    Debug.Print "Error " & Err.number & ": " & Err.Description
    Debug.Print "XML Preview: " & Left(xmlString, 500) & "..."

    MsgBox "Error in SaveToWord:" & vbCrLf & _
           "Line: " & Erl & vbCrLf & _
           "Error " & Err.number & ": " & Err.Description, vbCritical
End Sub


Public Function LoadFromWord() As Boolean
    Dim doc As Document: Set doc = ThisDocument
    Dim part As Office.CustomXMLPart

    For Each part In doc.CustomXMLParts
        If InStr(1, part.xml, "urn:my-waste-namespace") > 0 Then
            Me.LoadFromXml part.xml
            LoadFromWord = True
            Exit Function
        End If
    Next part

    LoadFromWord = False
End Function

Public Function Validate() As Boolean
    Dim w As Waste, ok As Boolean
    ok = True
    For Each w In wastes
        If w.id = "" Then
            Debug.Print "[ERROR] Waste has no ID."
            ok = False
        ElseIf w.WasteType = "" Then
            Debug.Print "[ERROR] Waste '" & w.id & "' has no type."
            ok = False
        End If
    Next w
    Validate = ok
End Function

Public Sub PruneOrphaned()
    Dim i As Long
    For i = wastes.count To 1 Step -1
        Dim w As Waste
        Set w = wastes.item(i)
        
        Debug.Print "[PRUNE] Checking waste: " & w.id
        
        Dim cc As ContentControl
        Set cc = ContentControlHelpers.FindControlByTag(w.id)
        
        ' === Check 1: Orphaned ===
        If cc Is Nothing Then
            Debug.Print "[PRUNE] ? Orphaned (no content control found)"
            wastes.Remove w.id
        
        ' === Check 2: Not inside ProcessDescription ===
        ElseIf Not ParentContentControlTitled("ProcessDescription", cc) Then
            Debug.Print "[PRUNE] ? Misplaced (not in ProcessDescription)"
            wastes.Remove w.id
        
        ' === (Optional) Check 3: Invalid ID Format ===
            Debug.Print "[PRUNE] ? OK: " & cc.Range.text
        End If
    Next i
End Sub

Public Sub SaveClean()
    Me.PruneOrphaned
    Me.SortByStepNumber
        'Me.SaveToWord
    Me.ValidateRoundTripWithCheck
End Sub

Public Sub ScaleAll(ByVal factor As Double)
    Dim i As Long
    For i = 1 To Me.count
        Dim w As Waste
        Set w = Me.item(i)
        If w.mass > 0 Then w.mass = w.mass * factor
        If w.Volume > 0 Then w.Volume = w.Volume * factor
        'Debug.Print "[SCALE] Waste " & w.ID & ": mass ? " & w.Mass & " g, volume ? " & w.Volume & " mL"
    Next i
End Sub

Public Sub SortByStepNumber(Optional descending As Boolean = False)
    Dim sortedList() As Waste
    Dim w As Waste
    Dim i As Long, j As Long
    Dim temp As Waste
    
    ' If there are fewer than 2 wastes, no need to sort:
    If wastes.count < 2 Then
        Exit Sub
    End If

    ' Copy to array
    ReDim sortedList(1 To wastes.count)
    For i = 1 To wastes.count
        Set sortedList(i) = wastes.item(i)
    Next i

    ' Insertion sort
    For i = 2 To UBound(sortedList)
        Set temp = sortedList(i)
        j = i - 1
        Do While j > 0
            Dim stepJ As Long, stepTemp As Long
            stepJ = GetStepNumberFromId(sortedList(j).id)
            stepTemp = GetStepNumberFromId(temp.id)

            If descending Then
                If stepTemp > stepJ Then
                    Set sortedList(j + 1) = sortedList(j)
                    j = j - 1
                Else
                    Exit Do
                End If
            Else
                If stepTemp < stepJ Then
                    Set sortedList(j + 1) = sortedList(j)
                    j = j - 1
                Else
                    Exit Do
                End If
            End If
        Loop
        Set sortedList(j + 1) = temp
    Next i

    ' Rebuild collection
    Set wastes = New Collection
    For i = 1 To UBound(sortedList)
        wastes.Add sortedList(i), sortedList(i).id
    Next i
End Sub

Public Sub Clone(oldTag As String, newTag As String)
    On Error GoTo Handler

    Dim original As Waste
    Set original = Me.GetById(oldTag)
    
    If original Is Nothing Then
        Debug.Print "[Clone] Waste not found: " & oldTag
        Exit Sub
    End If

    Dim cloned As Waste
    Set cloned = New Waste

    cloned.id = newTag
    cloned.mass = original.mass
    cloned.Volume = original.Volume
    cloned.WasteType = original.WasteType
    cloned.Description = original.Description

    Me.Add cloned

    Debug.Print "[Clone] Waste cloned: " & oldTag & " ? " & newTag
    Exit Sub

Handler:
    Debug.Print "[ERROR] WasteCollection.Clone failed: " & Err.Description
    Err.Clear
End Sub

'=== WasteCollection: Validate Round Trip ===
Public Sub ValidateRoundTripWithCheck()
    Dim originalXml As String, reloadedXml As String
    Dim reloaded As New WasteCollection
    Dim w As Waste, reW As Waste
    Dim mismatchFound As Boolean: mismatchFound = False
    Dim errorDetails As String

    originalXml = Me.ToXml

    Me.SaveToWord
    If Not reloaded.LoadFromWord Then
        MsgBox "[RoundTrip] Reload failed: no XML part found.", vbCritical
        Exit Sub
    End If

    If Me.count <> reloaded.count Then
        errorDetails = errorDetails & "- Waste count mismatch: " & Me.count & " vs " & reloaded.count & vbCrLf
        mismatchFound = True
    End If

    Dim i As Long
    For i = 1 To Me.count
        Set w = Me.item(i)
        Set reW = reloaded.GetById(w.id)
        If reW Is Nothing Then
            errorDetails = errorDetails & "- Missing waste ID after reload: " & w.id & vbCrLf
            mismatchFound = True
        End If
    Next i

    reloadedXml = reloaded.ToXml
    If originalXml <> reloadedXml Then
        errorDetails = errorDetails & "- XML mismatch detected." & vbCrLf
        mismatchFound = True
    End If

    If mismatchFound Then
        Dim userMessage As String
        userMessage = "CRITICAL ERROR: RoundTrip validation failed for WASTE COLLECTION and may lead to DATA LOSS." & vbCrLf & vbCrLf & _
                      "Please take a PRINT SCREEN of this message and SEND the FILE to: klm" & vbCrLf & _
                      "Then: Save the file as a new version, reload it, and scroll to the 'Mass Balance Table'." & vbCrLf & _
                      "Use the macro ribbon to trigger table updates and check for lost compounds or waste entries." & vbCrLf & vbCrLf & _
                      "--- DETAILS ---" & vbCrLf & errorDetails

        MsgBox userMessage, vbCritical + vbOKOnly, "[WasteRoundTrip] Validation Error"
    End If
End Sub




