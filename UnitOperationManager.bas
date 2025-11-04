Attribute VB_Name = "UnitOperationManager"
' UnitOperationManager Module
Option Explicit

' Define Word constants for fallback
Const wdContentControlDropdown As Long = 3


' Helper function to check if two tables are the same
Public Function AreTablesEqual(table1 As Table, table2 As Table) As Boolean
    AreTablesEqual = (table1.Range.Start = table2.Range.Start And table1.Range.End = table2.Range.End)
End Function

Sub HandleDropdownSelection(ByVal ccDropdown As ContentControl)
    On Error GoTo ErrorHandler
    ' Debug.Print "Starting HandleDropdownSelection..."

    Dim selectedEntry As String
    Dim selectedKey As String
    Dim sourceTable As Table
    Dim sourceRow As row
    Dim targetRow As row
    Dim ccUnitOperation As ContentControl
    Dim dropdownEntry As ContentControlListEntry

    ' Get the selected entry from the dropdown
    selectedEntry = Trim(ccDropdown.Range.text)
    If selectedEntry = "" Then
        Debug.Print "No selection made in the dropdown."
        Exit Sub
    End If

    'Debug.Print "Selected entry: " & selectedEntry

    ' Extract the key (Value) of the selected dropdown entry
    selectedKey = ""
    For Each dropdownEntry In ccDropdown.DropdownListEntries
        If dropdownEntry.text = selectedEntry Then
            selectedKey = dropdownEntry.value
            Exit For
        End If
    Next dropdownEntry

    If selectedKey = "" Then
        Debug.Print "No key found for the selected entry."
        Exit Sub
    Else
        ' Debug.Print "Selected key: " & selectedKey
    End If

    ' Get the source table via its content control
    Set sourceTable = GetTableFromContentControl("DefinitionOfUnitOperations")
    If sourceTable Is Nothing Then
        Debug.Print "Source table (DefinitionOfUnitOperations) not found."
        Exit Sub
    Else
        ' Debug.Print "Source table located successfully."
    End If

    ' Search for the matching row in the source table
    For Each sourceRow In sourceTable.rows
        ' Extract the content control from the first column
        Set ccUnitOperation = FetchContentControlFromCell(sourceRow.Cells(1))
        If Not ccUnitOperation Is Nothing Then
            If Trim(ccUnitOperation.title) = selectedKey Then
                ' Debug.Print "Match found for key: " & selectedKey
                ' Replace the target row with the source row content
                Set targetRow = ccDropdown.Range.rows(1)
                sourceRow.Range.Copy
                targetRow.Range.PasteAndFormat wdFormatOriginalFormatting
                ' Debug.Print "Row replaced with the selected unit operation."

                ' Call UnitOperationInitializer to assign unique tags
                Call UnitOperationHandler.InitializeUnitOperation(targetRow)
                Exit Sub
            End If
        Else
            Debug.Print "No content control found in the first cell of row " & sourceRow.index
        End If
    Next sourceRow

    Debug.Print "No matching unit operation found for the selected key."
    Exit Sub

ErrorHandler:
    Debug.Print "Error in HandleDropdownSelection: " & Err.Description
    Err.Clear
End Sub




'----- Move to COntentControl Helpers or Table Helpers ----' Consider renaming helper modules to Helper_ContentControls
' Get the table contained within a specified rich text content control
Private Function GetTableFromContentControl(ByVal controlTitle As String) As Table
   ' Debug.Print "Attempting to locate content control: " & controlTitle
    Dim cc As ContentControl

    ' Locate the content control by title
    For Each cc In ThisDocument.contentControls
        If cc.title = controlTitle Then
            If cc.Range.Tables.count > 0 Then
                Set GetTableFromContentControl = cc.Range.Tables(1)
               ' Debug.Print "Table located in content control: " & controlTitle
                Exit Function
            End If
        End If
    Next cc

    Debug.Print "No table found in content control: " & controlTitle
    Set GetTableFromContentControl = Nothing
End Function




