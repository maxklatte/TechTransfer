Attribute VB_Name = "CloneHelper"
Option Explicit

'Can be deleted.
Public Sub CloneCollection(sourceCol As Collection, targetCol As Collection)
    Dim item As Variant
    Dim clonedDict As Object
    Dim key As Variant

    For Each item In sourceCol
        If TypeName(item) = "Dictionary" Then
            Set clonedDict = CreateObject("Scripting.Dictionary")
            For Each key In item.keys
                clonedDict.Add key, item(key)
            Next key
            targetCol.Add clonedDict
        End If
    Next item
End Sub

' === Enhanced Helper for Collection Cloning ===
Public Sub CloneCollectionWithNewTags(sourceCol As Collection, targetCol As Collection, baseID As String, collectionType As String, deepClone As Boolean)
    Dim item As Variant
    Dim clonedDict As Object
    Dim key As Variant
    Dim oldTag As String
    Dim newTag As String

    For Each item In sourceCol
        If TypeName(item) = "Dictionary" Then
            Set clonedDict = CreateObject("Scripting.Dictionary")

            For Each key In item.keys
                If key = "Tag" Then
                    oldTag = item(key)
                    newTag = baseID & Right(oldTag, 2)
                    clonedDict.Add key, newTag
                    Debug.Print "[Clone] Tag updated: old: " & oldTag & "   new: " & newTag
                    If deepClone Then
                        CloneAssociatedObject oldTag, newTag, collectionType
                    End If
                Else
                    clonedDict.Add key, item(key)
                End If
            Next key

            targetCol.Add clonedDict
        End If
    Next item
End Sub

Private Sub CloneAssociatedObject(oldTag As String, newTag As String, collectionType As String)
    On Error GoTo Handler

    Select Case LCase(collectionType)
        Case "inputs"
            Debug.Print "[Clone] Trigger CompoundCollection clone: " & oldTag & " ? " & newTag
            GetGlobalCompoundCollection().Clone oldTag, newTag

        Case "outputs"
            Debug.Print "[Clone] Trigger WasteCollection clone: " & oldTag & " ? " & newTag
            GetGlobalWasteCollection().Clone oldTag, newTag

        Case Else
            Debug.Print "[Clone] No associated collection for: " & collectionType & " [tag: " & oldTag & "]"
    End Select
    Exit Sub

Handler:
    Debug.Print "[ERROR] CloneAssociatedObject failed: " & Err.Description & " [" & collectionType & "]"
    Err.Clear
End Sub


