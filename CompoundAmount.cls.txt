VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CompoundAmount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=== FILE: CompoundAmount.cls ===
Option Explicit

'=== Private storage ===
Private pMass As Double
Private pVolume As Double
Private pMolarAmount As Double
Private pAssay As Double ' 0-100%

'=== Public Properties ===
Public Property Get mass() As Double
    mass = pMass
End Property

Public Property Get Volume() As Double
    Volume = pVolume
End Property

Public Property Get MolarAmount() As Double
    MolarAmount = pMolarAmount
End Property

Public Property Get Assay() As Double
    If pAssay <= 0 Then
        Assay = 100#
    Else
        Assay = pAssay
    End If
End Property

Public Property Let Assay(ByVal value As Double)
    If value < 0 Or value > 100 Then
        Err.Raise vbObjectError + 5000, "CompoundAmount", "Assay must be between 0 and 100."
    End If
    pAssay = value
End Property

'=== Stateless Calculations ===
Public Property Get correctedMass() As Double
    correctedMass = Me.mass * (Me.Assay / 100#)
End Property

Public Function GetCorrectedMolarAmount(ByVal stoff As Stoffdaten) As Double
    ValidateStoffdaten stoff
    GetCorrectedMolarAmount = Me.correctedMass / stoff.molarMass
End Function

Public Function GetEquiv(ByVal reference As CompoundAmount, _
                         ByVal selfStoff As Stoffdaten, _
                         ByVal refStoff As Stoffdaten) As Double
    Dim selfMol As Double
    Dim refMol As Double

    selfMol = Me.correctedMass / selfStoff.molarMass
    refMol = reference.correctedMass / refStoff.molarMass

    If refMol <= 0 Then
        Err.Raise vbObjectError + 5101, "CompoundAmount", "Reference corrected molar amount must be greater than zero."
    End If

    GetEquiv = selfMol / refMol
End Function

Public Function GetRelativeVolume(ByVal reference As CompoundAmount) As Double
    ValidateReference reference

    Dim refMass As Double
    refMass = reference.correctedMass

    If refMass <= 0 Then
        Err.Raise vbObjectError + 5102, "CompoundAmount", "Reference corrected mass must be greater than zero."
    End If

    GetRelativeVolume = Me.Volume / refMass
End Function

'=== Update Mechanisms ===
Public Sub SetMass(ByVal newMass As Double, ByVal stoff As Stoffdaten)
    ValidateStoffdaten stoff
    If newMass < 0 Then Err.Raise vbObjectError + 5001, "CompoundAmount", "Mass must be non-negative."
    
    pMass = newMass
    pMolarAmount = pMass / stoff.molarMass
    pVolume = pMass / stoff.Density
End Sub

Public Sub SetMolarAmount(ByVal newMolarAmount As Double, ByVal stoff As Stoffdaten)
    ValidateStoffdaten stoff
    If newMolarAmount < 0 Then Err.Raise vbObjectError + 5002, "CompoundAmount", "Molar amount must be non-negative."
    
    pMolarAmount = newMolarAmount
    pMass = pMolarAmount * stoff.molarMass
    pVolume = pMass / stoff.Density
End Sub

Public Sub SetVolume(ByVal newVolume As Double, ByVal stoff As Stoffdaten)
    ValidateStoffdaten stoff
    If newVolume < 0 Then Err.Raise vbObjectError + 5003, "CompoundAmount", "Volume must be non-negative."

    pVolume = newVolume
    pMass = pVolume * stoff.Density
    pMolarAmount = pMass / stoff.molarMass
End Sub

Public Sub SetEquiv(ByVal equivValue As Double, _
                    ByVal reference As CompoundAmount, _
                    ByVal selfStoff As Stoffdaten, _
                    ByVal refStoff As Stoffdaten)
    
    ValidateReference reference
    ValidateStoffdaten selfStoff
    ValidateStoffdaten refStoff

    If equivValue < 0 Then
        Err.Raise vbObjectError + 5201, "CompoundAmount", "Equivalence value must be non-negative."
    End If

    Dim refCorrectedMass As Double
    refCorrectedMass = reference.correctedMass

    If refCorrectedMass <= 0 Then
        Err.Raise vbObjectError + 5101, "CompoundAmount", "Reference corrected mass must be greater than zero."
    End If

    Dim refMol As Double
    refMol = refCorrectedMass / refStoff.molarMass

    If refMol <= 0 Then
        Err.Raise vbObjectError + 5102, "CompoundAmount", "Reference molar amount must be greater than zero."
    End If

    Dim targetCorrectedMol As Double
    targetCorrectedMol = equivValue * refMol

    Dim targetCorrectedMass As Double
    targetCorrectedMass = targetCorrectedMol * selfStoff.molarMass

    Dim targetRawMass As Double
    targetRawMass = targetCorrectedMass * (100# / Me.Assay)

    Me.SetMass targetRawMass, selfStoff
End Sub

Public Sub SetRelativeVolume(ByVal relValue As Double, ByVal reference As CompoundAmount, ByVal stoff As Stoffdaten)
    ValidateReference reference
    ValidateStoffdaten stoff

    If relValue < 0 Then
        Err.Raise vbObjectError + 5202, "CompoundAmount", "Relative volume value must be non-negative."
    End If

    Dim refMass As Double
    refMass = reference.correctedMass

    If refMass <= 0 Then
        Err.Raise vbObjectError + 5102, "CompoundAmount", "Reference corrected mass must be greater than zero."
    End If

    Dim targetVolume As Double
    targetVolume = relValue * refMass

    Me.SetVolume targetVolume, stoff
End Sub

Public Sub SetCorrectedMass(ByVal targetCorrectedMass As Double, ByVal stoff As Stoffdaten)
    ValidateStoffdaten stoff
    If targetCorrectedMass < 0 Then
        Err.Raise vbObjectError + 5301, "CompoundAmount", "Corrected mass must be non-negative."
    End If

    Dim rawMass As Double
    rawMass = targetCorrectedMass * (100# / Me.Assay)
    Me.SetMass rawMass, stoff
End Sub

Public Sub SetCorrectedMolarAmount(ByVal targetCorrectedMol As Double, ByVal stoff As Stoffdaten)
    ValidateStoffdaten stoff
    If targetCorrectedMol < 0 Then
        Err.Raise vbObjectError + 5302, "CompoundAmount", "Corrected molar amount must be non-negative."
    End If

    Dim correctedMass As Double
    correctedMass = targetCorrectedMol * stoff.molarMass

    Dim rawMass As Double
    rawMass = correctedMass * (100# / Me.Assay)
    Me.SetMass rawMass, stoff
End Sub

'=== Internal Validation ===
Private Sub ValidateStoffdaten(ByVal stoff As Stoffdaten)
    If stoff Is Nothing Then Err.Raise vbObjectError + 5004, "CompoundAmount", "Stoffdaten is required."
    
    'Added quick fix to deal with not defined molar masses, may be adressed differntly in future revisions
If stoff.Density <= 0 And stoff.Density <> 1 Then
    MsgBox "Warning: Density is zero or undefined for compound. It will be treated as 1.", vbExclamation, "Invalid Density"
    stoff.Density = 1
End If
    

   If stoff.molarMass <= 0 And stoff.molarMass <> -1 Then
    MsgBox "Warning: Molar mass is zero or undefined for compound. It will be treated as -1.", vbExclamation, "Invalid Molar Mass"
    stoff.molarMass = -1
    
End If
End Sub

Private Sub ValidateReference(ByVal reference As CompoundAmount)
    If reference Is Nothing Then
        Err.Raise vbObjectError + 5007, "CompoundAmount", "Reference CompoundAmount is required."
    End If
End Sub
