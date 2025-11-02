VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Stoffdaten"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pProductCode As String
Private pTitle As String
Private pCASReg As String
Private pDensity As Double
Private pMolarMass As Double

Public Property Get productCode() As String: productCode = pProductCode: End Property
Public Property Let productCode(ByVal value As String): pProductCode = value: End Property

Public Property Get title() As String: title = pTitle: End Property
Public Property Let title(ByVal value As String): pTitle = value: End Property

Public Property Get CASReg() As String: CASReg = pCASReg: End Property
Public Property Let CASReg(ByVal value As String): pCASReg = value: End Property

Public Property Get Density() As Double: Density = pDensity: End Property
Public Property Let Density(ByVal value As Double): pDensity = value: End Property

Public Property Get molarMass() As Double: molarMass = pMolarMass: End Property
Public Property Let molarMass(ByVal value As Double): pMolarMass = value: End Property

'=== Updated Stoffdaten.DisplayInfo ===
Public Sub DisplayInfo()
    Dim info As String
    info = "Product Code: " & pProductCode & vbCrLf & _
           "Title: " & pTitle & vbCrLf & _
           "CASReg: " & pCASReg & vbCrLf & _
           "Density [g/mL]: " & pDensity & vbCrLf & _
           "Molar Mass [g/mol]: " & pMolarMass
    MsgBox info, vbInformation, "Stoffdaten Info"
End Sub

