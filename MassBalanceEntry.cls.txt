VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MassBalanceEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Step As Long
Public UnitOperation As String
Public title As String
Public Weight As Double
Public Volume As Double
Public EntryType As String ' "compound" or "waste"
Public id As String

