VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAppEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' === clsAppEvents ===
Option Explicit

Public WithEvents appWord As word.Application
Attribute appWord.VB_VarHelpID = -1

Private Sub appWord_DocumentBeforeSave( _
    ByVal doc As Document, _
    SaveAsUI As Boolean, _
    Cancel As Boolean)

    On Error GoTo ErrorHandler

    Debug.Print "[Save] Save Compound and Waste Collection"
    
    GetGlobalCompoundCollection().SaveClean
    GetGlobalWasteCollection().SaveClean

    Exit Sub

ErrorHandler:
    Debug.Print "!! Error during SaveClean: " & Err.Description
    ' Allow save to proceed regardless of SaveClean failure
    Resume Next
End Sub

