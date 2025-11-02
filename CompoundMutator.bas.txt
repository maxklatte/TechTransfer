Attribute VB_Name = "CompoundMutator"
'=== MODULE: CompoundMutator (UPDATED with Identification Change Detection) ===
' Purpose: Unified module handling parsing, diffing, conflict detection, identification change detection, and applying edits.

Option Explicit

'=== PUBLIC MAIN ENTRY POINT ===
' Parses edit string, diffs, checks identification change/conflicts, applies if safe
Public Sub ApplyEditString(ByVal cmp As Compound, ByVal editStr As String, ByVal reference As Compound)
    Dim editDict As Object
    Dim changedParts As Collection

    Set editDict = ParseEditString(editStr)
    Set changedParts = GetChangedParts(cmp, editDict, reference)

    If IsIdentificationChanged(changedParts) Then
        Debug.Print "[INFO] Detected identification change (product-code or title)."
        ' Proceed to apply changes even if multiple identification fields changed
        ApplyChanges cmp, changedParts, reference
        Exit Sub
    End If

    If IsEditConflicting(changedParts) Then
      MsgBox "Please modify only one of: mass, volume, amount, equiv, or rel.volume.", vbExclamation, "Conflict Detected"
        Err.Raise vbObjectError + 4001, "CompoundMutator.ApplyEditString", _
                  "Conflict detected: Multiple quantity fields changed simultaneously."
    End If

    ApplyChanges cmp, changedParts, reference
End Sub


'=== INTERNAL HELPERS ===

Private Function ParseEditString(ByVal editStr As String) As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim parts() As String: parts = Split(editStr, ";")
    Dim i As Long, kv() As String
    Dim rawKey As String, normKey As String, value As String

    For i = LBound(parts) To UBound(parts)
        kv = Split(parts(i), "=")
        If UBound(kv) <> 1 Then GoTo Skip

        rawKey = Trim(kv(0))
        value = Trim(kv(1))
        normKey = NormalizeEditKey(rawKey)

        If Not dict.Exists(normKey) Then
            dict.Add normKey, value
        End If
Skip:
    Next i

    Set ParseEditString = dict
End Function


Private Function GetChangedParts(ByVal cmp As Compound, ByVal editDict As Object, ByVal reference As Compound) As Collection
    Dim result As New Collection
    Dim currentDict As Object: Set currentDict = CreateObject("Scripting.Dictionary")
    Dim i As Long
    Dim parts() As String
    Dim key As Variant
    Dim displayString As String

    Debug.Print "[DBG] === ToDisplayString(edit) Output ==="

    On Error Resume Next
    displayString = cmp.ToDisplayString("edit", reference)
    If Err.number <> 0 Then
        Debug.Print "[ERR] ToDisplayString failed: " & Err.Description
        displayString = ""
        Err.Clear
    End If
    On Error GoTo 0

    If Len(displayString) = 0 Then
        Debug.Print "[FATAL] Display string is empty. Cannot parse fields."
        Set GetChangedParts = result
        Exit Function
    End If

    parts = Split(displayString, ";")

    For i = LBound(parts) To UBound(parts)
        Dim kv() As String
        kv = Split(Trim(parts(i)), "=")
        If UBound(kv) = 1 Then
            currentDict(NormalizeEditKey(kv(0))) = Trim(kv(1))
        End If
    Next i

    For Each key In editDict.keys
        If Not currentDict.Exists(key) Then
            Debug.Print "[DBG] CHANGED (missing): " & key & "=" & editDict(key)
            result.Add key & "=" & editDict(key)
        ElseIf currentDict(key) <> editDict(key) Then
            Debug.Print "[DBG] CHANGED (diff val): " & key & ": [" & currentDict(key) & "] -> [" & editDict(key) & "]"
            result.Add key & "=" & editDict(key)
        End If
    Next key

    Debug.Print "[DBG] === Final Changed Parts ==="
    For Each key In result
        Debug.Print "  " & key
    Next

    Set GetChangedParts = result
End Function



Private Function IsIdentificationChanged(ByVal changedParts As Collection) As Boolean
    Dim i As Long
    Dim kv() As String
    Dim key As String

    For i = 1 To changedParts.count
        kv = Split(changedParts(i), "=")
        If UBound(kv) <> 1 Then GoTo Skip

        key = NormalizeEditKey(Trim(kv(0)))

        Select Case key
            Case "product-code", "title"
                IsIdentificationChanged = True
                Exit Function
        End Select
Skip:
    Next i

    IsIdentificationChanged = False
End Function


Private Function IsEditConflicting(ByVal changedParts As Collection) As Boolean
    Dim qtyFields As Object
    Set qtyFields = CreateObject("Scripting.Dictionary")

    qtyFields.Add "mass", True
    qtyFields.Add "volume", True
    qtyFields.Add "amount", True
    qtyFields.Add "equiv", True
    qtyFields.Add "rel.volume", True

    Dim conflictCount As Long: conflictCount = 0
    Dim i As Long
    Dim kv() As String
    Dim key As String

    For i = 1 To changedParts.count
        kv = Split(changedParts(i), "=")
        If UBound(kv) <> 1 Then GoTo Skip

        key = NormalizeEditKey(Trim(kv(0)))

        If qtyFields.Exists(key) Then
            conflictCount = conflictCount + 1
        End If
Skip:
    Next i

    IsEditConflicting = (conflictCount > 1)
End Function







Private Sub ApplyAssayChange(ByVal cmp As Compound, ByVal newAssay As Double)
    ' Step 1: Snapshot current state (unmutated)
    Dim oldMass As Double: oldMass = cmp.amount.mass
    Dim oldCorrectedMass As Double: oldCorrectedMass = cmp.amount.correctedMass
    Dim oldMol As Double: oldMol = cmp.amount.GetCorrectedMolarAmount(cmp.Stoffdaten) * 1000

    ' Step 2: Simulate new values without mutating cmp
    Dim molarMass As Double: molarMass = cmp.Stoffdaten.molarMass
    Dim projectedNewMass As Double: projectedNewMass = oldCorrectedMass * (100# / newAssay)
    Dim projectedNewCorrectedMass As Double: projectedNewCorrectedMass = oldMass * newAssay / 100#
    Dim projectedNewMol As Double: projectedNewMol = projectedNewCorrectedMass / molarMass * 1000

    ' Step 3: Print debug info for verification
    'Debug.Print "[DEBUG] Assay Change Detected"
    'Debug.Print "  Current Assay: " & cmp.amount.Assay
    'Debug.Print "  New Assay: " & newAssay
    'Debug.Print "  Old Mass: " & format(oldMass, "0.00") & " g"
    'Debug.Print "  Old Corrected Molar Amount: " & format(oldMol, "0.0") & " mmol"
    'Debug.Print "  Projected New Mass: " & format(projectedNewMass, "0.00") & " g"
    'Debug.Print "  Projected New Molar Amount: " & format(projectedNewMol, "0.0") & " mmol"

    ' Step 4: Prompt user with simulated projections
    Dim prompt As String
    prompt = "Assay changed." & vbCrLf & _
             "Adjust mass to keep molar amount [mmol] unchanged?" & vbCrLf & vbCrLf & _
             "Yes = Adjust mass from " & format(oldMass, "0.00") & " g to " & format(projectedNewMass, "0.00") & " g." & vbCrLf & _
             "No = Adjust molar amount from " & format(oldMol, "0.0") & " mmol to " & format(projectedNewMol, "0.0") & " mmol." & vbCrLf & _
             "Cancel = Abort."

    Dim choice As VbMsgBoxResult
    choice = MsgBox(prompt, vbYesNoCancel + vbQuestion, "Assay Adjustment")

    ' Step 5: Apply mutation based on choice
    If choice = vbYes Then
        cmp.amount.Assay = newAssay
        cmp.amount.SetCorrectedMass oldCorrectedMass, cmp.Stoffdaten
        Debug.Print "[ACTION] Mass adjusted to maintain corrected molar amount."
    ElseIf choice = vbNo Then
        cmp.amount.Assay = newAssay
        Debug.Print "[ACTION] Mass retained. Corrected molar amount will change."
    Else
        Debug.Print "[CANCELLED] Assay edit aborted by user."
        Exit Sub
    End If
End Sub




'=== UpdateProductCode: Looks up Stoffdaten by product code and applies it ===
Public Sub UpdateProductCode(ByVal cmp As Compound, ByVal newKey As String)
    Debug.Print "[STEP 1] Starting UpdateProductCode with key: " & newKey

    Dim stoffList As Collection
    Set stoffList = GetGlobalStoffdatenbank().FindMatchingStoffe(newKey)

    If stoffList.count = 0 Then
        Debug.Print "[WARN] No matching Stoffdaten found for: " & newKey
        Exit Sub
    End If

    Dim selectedStoff As Stoffdaten
    If stoffList.count = 1 Then
        Set selectedStoff = stoffList(1)
        Debug.Print "[INFO] One Stoffdaten match found. Auto-selected."
    Else
        Debug.Print "[INFO] Multiple Stoffdaten matches found. Prompting user."
        Dim choice As Long
        choice = GetGlobalStoffdatenbank().PromptStoffIndexSelection(stoffList)
        If choice = 0 Then
            Debug.Print "[INFO] User cancelled Stoffdaten selection."
            Exit Sub
        End If
        Set selectedStoff = stoffList(choice)
    End If
    
   Debug.Print "MOLAR MASS: " & selectedStoff.molarMass
   Debug.Print "DENSITY: " & selectedStoff.Density
    
    ' Guard against invalid molarMass
    If selectedStoff.molarMass = 0 Then
         MsgBox "Warning: Stoffdaten for product '" & selectedStoff.productCode & "' has molar mass = 0. It will be treated as -1.", vbExclamation, "Invalid Molar Mass"
         selectedStoff.molarMass = -1
    End If

    ' Optional: Guard against invalid density
    If selectedStoff.Density = 0 Then
        MsgBox "Warning: Stoffdaten for product '" & selectedStoff.productCode & "' has density = 0. It will be treated as -1.", vbExclamation, "Invalid Density"
        selectedStoff.Density = 1
    End If



    Debug.Print "[STEP 2] Updating Stoffdaten in Compound to: " & selectedStoff.productCode
    Set cmp.Stoffdaten = selectedStoff

    
    ' Only infer a new type if it isn’t already “product”
If cmp.compoundType <> "product" Then
    cmp.compoundType = InferCompoundTypeFromCode(selectedStoff.productCode)
End If
    cmp.InstructAssayCorrection = False


    Debug.Print "[STEP 3] Triggering recalculation based on new Stoffdaten."
    cmp.amount.SetMass cmp.amount.mass, cmp.Stoffdaten
End Sub

Private Function NormalizeEditKey(ByVal k As String) As String
    On Error GoTo HandleError

    k = LCase(Trim(k))

    ' Prefer exact keywords over substring matches
    If k = "instruct assay correction" Or k = "instructassaycorrection" Then
        NormalizeEditKey = "instructAssayCorrection"
        Exit Function
    End If

    If InStr(k, "correct for assay") > 0 Or InStr(k, "instruct assay") > 0 Then
        NormalizeEditKey = "instructAssayCorrection"
        Exit Function
    End If

    If InStr(k, "product-code") > 0 Then
        NormalizeEditKey = "product-code"
        Exit Function
    End If

    If InStr(k, "title") > 0 Then
        NormalizeEditKey = "title"
        Exit Function
    End If

    If InStr(k, "mass") > 0 And InStr(k, "corrected") = 0 Then
        NormalizeEditKey = "mass"
        Exit Function
    End If

    If InStr(k, "volume") > 0 And InStr(k, "rel") = 0 Then
        NormalizeEditKey = "volume"
        Exit Function
    End If

    If InStr(k, "amount") > 0 And InStr(k, "corrected") = 0 Then
        NormalizeEditKey = "amount"
        Exit Function
    End If

    If InStr(k, "equiv") > 0 Then
        NormalizeEditKey = "equiv"
        Exit Function
    End If

    If InStr(k, "rel.volume") > 0 Or InStr(k, "relative-volume") > 0 Then
        NormalizeEditKey = "rel.volume"
        Exit Function
    End If

    If InStr(k, "assay") > 0 Then
        NormalizeEditKey = "assay"
        Exit Function
    End If

    If InStr(k, "type") > 0 Then
        NormalizeEditKey = "compoundType"
        Exit Function
    End If

    If InStr(k, "corrected amount") > 0 Then
        NormalizeEditKey = "corrected amount [g]"
        Exit Function
    End If

    NormalizeEditKey = k ' fallback
    Exit Function

HandleError:
    NormalizeEditKey = k
End Function

Private Function ParseBoolean(ByVal val As String) As Boolean
    Dim norm As String
    norm = LCase(Trim(val))

    Select Case norm
        Case "true", "yes", "wahr"
            ParseBoolean = True
        Case "false", "no", "falsch"
            ParseBoolean = False
        Case Else
            Err.Raise vbObjectError + 5500, "ParseBoolean", "Invalid boolean value: " & val
    End Select
End Function
Private Function GetAttributeGroup(ByVal key As String) As String
    Select Case key
        Case "product-code", "title"
            GetAttributeGroup = "identity"
        Case "mass", "volume", "amount", "equiv", "rel.volume", "corrected amount [g]"
            GetAttributeGroup = "quantity"
        Case "assay"
            GetAttributeGroup = "assay"
        Case "compoundType", "instructAssayCorrection"
            GetAttributeGroup = "flags"
        Case Else
            GetAttributeGroup = "unknown"
    End Select
End Function

Private Function DetectGroupConflicts(ByVal changedParts As Collection) As Boolean
    Dim groupCounts As Object: Set groupCounts = CreateObject("Scripting.Dictionary")
    Dim key As String, group As String
    Dim i As Long, kv() As String

    For i = 1 To changedParts.count
        kv = Split(changedParts(i), "=")
        If UBound(kv) <> 1 Then GoTo Skip

        key = NormalizeEditKey(Trim(kv(0)))
        group = GetAttributeGroup(key)

        ' Track count per group
        If group <> "unknown" Then
            If Not groupCounts.Exists(group) Then
                groupCounts.Add group, 1
            Else
                groupCounts(group) = groupCounts(group) + 1
            End If
        End If
Skip:
    Next i

    ' Special case: allow mass+assay OR corrected+assay
    If groupCounts.Exists("assay") And groupCounts("assay") = 1 Then
        If groupCounts.Exists("quantity") And groupCounts("quantity") = 1 Then
            ' Only allow if exact keys match one of the known valid combos
            Dim keys As Collection: Set keys = ExtractNormalizedKeys(changedParts)
            Dim required1 As Variant: required1 = Array("mass", "assay")
            Dim required2 As Variant: required2 = Array("corrected amount [g]", "assay")
            If HasKeys(keys, required1) Or HasKeys(keys, required2) Then
                DetectGroupConflicts = False
                Exit Function
            End If
        End If
    End If

    ' Quantity or assay group must not have >1 change (outside special case)
    If groupCounts.Exists("quantity") And groupCounts("quantity") > 1 Then
        DetectGroupConflicts = True
        Exit Function
    End If
    If groupCounts.Exists("assay") And groupCounts("assay") > 1 Then
        DetectGroupConflicts = True
        Exit Function
    End If

    DetectGroupConflicts = False
End Function

Private Function ExtractNormalizedKeys(changedParts As Collection) As Collection
    Dim keys As New Collection
    Dim i As Long, kv() As String, key As String

    For i = 1 To changedParts.count
        kv = Split(changedParts(i), "=")
        If UBound(kv) = 1 Then
            key = NormalizeEditKey(Trim(kv(0)))
            keys.Add key
        End If
    Next i

    Set ExtractNormalizedKeys = keys
End Function

Private Function HasKeys(keys As Collection, required As Variant) As Boolean
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim i As Long

    For i = 1 To keys.count
        dict(keys(i)) = True
    Next i

    For i = LBound(required) To UBound(required)
        If Not dict.Exists(required(i)) Then
            HasKeys = False
            Exit Function
        End If
    Next i

    HasKeys = True
End Function

Private Sub ApplyChanges(ByVal cmp As Compound, ByVal changedParts As Collection, ByVal reference As Compound)
    Debug.Print "[DBG] === ApplyChanges Start ==="

    ApplyGroupChanges cmp, changedParts, reference, "identity"
    ApplyGroupChanges cmp, changedParts, reference, "quantity"
    ApplyGroupChanges cmp, changedParts, reference, "assay"

    RecalculateMassOrCorrectedAmount cmp, changedParts

    ApplyGroupChanges cmp, changedParts, reference, "flags"

    Debug.Print "[DBG] === ApplyChanges End ==="
End Sub

Private Sub RecalculateMassOrCorrectedAmount(ByVal cmp As Compound, ByVal changedParts As Collection)
    Dim keys As Collection: Set keys = ExtractNormalizedKeys(changedParts)

    If HasKeys(keys, Array("mass", "assay")) Then
        Debug.Print "[DBG] Special case detected: mass + assay — recalculating corrected mass"
        'Debug.Print "[DBG] BEFORE SetMass: mass=" & cmp.amount.Mass & ", assay=" & cmp.amount.Assay & ", corrected=" & cmp.amount.correctedMass
        cmp.amount.SetMass cmp.amount.mass, cmp.Stoffdaten
        'Debug.Print "[DBG] AFTER SetMass: mass=" & cmp.amount.Mass & ", assay=" & cmp.amount.Assay & ", corrected=" & cmp.amount.correctedMass

    ElseIf HasKeys(keys, Array("corrected amount [g]", "assay")) Then
        Debug.Print "[DBG] Special case detected: corrected amount + assay — recalculating total mass"
        Dim userCorrectedValue As Double
        userCorrectedValue = CDbl(GetChangedValue(changedParts, "corrected amount [g]"))
        'Debug.Print "[DBG] USING corrected amount from input: " & userCorrectedValue
        'Debug.Print "[DBG] BEFORE SetCorrectedMass: mass=" & cmp.amount.Mass & ", assay=" & cmp.amount.Assay & ", corrected=" & cmp.amount.correctedMass
        cmp.amount.SetCorrectedMass userCorrectedValue, cmp.Stoffdaten
        'Debug.Print "[DBG] AFTER SetCorrectedMass: mass=" & cmp.amount.Mass & ", assay=" & cmp.amount.Assay & ", corrected=" & cmp.amount.correctedMass
    End If
End Sub

Private Function GetChangedValue(ByVal changedParts As Collection, ByVal key As String) As String
    Dim i As Long, kv() As String, norm As String
    For i = 1 To changedParts.count
        kv = Split(changedParts(i), "=")
        If UBound(kv) = 1 Then
            norm = NormalizeEditKey(Trim(kv(0)))
            If norm = key Then
                GetChangedValue = Trim(kv(1))
                Exit Function
            End If
        End If
    Next i
    GetChangedValue = ""
End Function

Private Sub ApplyGroupChanges(ByVal cmp As Compound, ByVal changedParts As Collection, ByVal reference As Compound, ByVal group As String)
    Dim i As Long, kv() As String, key As String, value As String
    Dim onlyAssayChange As Boolean: onlyAssayChange = (changedParts.count = 1)

    For i = 1 To changedParts.count
        kv = Split(changedParts(i), "=")
        If UBound(kv) <> 1 Then GoTo SkipApply

        key = NormalizeEditKey(Trim(kv(0)))
        value = Trim(kv(1))

        If GetAttributeGroup(key) <> group Then GoTo SkipApply

        Debug.Print "[DBG] Applying " & key & " = " & value & " in group: " & group

        Select Case key
            Case "product-code", "title"
                CompoundMutator.UpdateProductCode cmp, value

            Case "mass"
                cmp.amount.SetMass CDbl(value), cmp.Stoffdaten

            Case "volume"
                cmp.amount.SetVolume CDbl(value), cmp.Stoffdaten

            Case "amount"
                cmp.amount.SetCorrectedMolarAmount CDbl(value) / 1000#, cmp.Stoffdaten

            Case "equiv"
                cmp.amount.SetEquiv CDbl(value), reference.amount, cmp.Stoffdaten, reference.Stoffdaten

            Case "rel.volume"
                cmp.amount.SetRelativeVolume CDbl(value), reference.amount, cmp.Stoffdaten

            Case "assay"
                If onlyAssayChange Then
                    Call ApplyAssayChange(cmp, CDbl(value))
                Else
                    cmp.amount.Assay = CDbl(value)
                End If

            Case "instructAssayCorrection"
                cmp.InstructAssayCorrection = ParseBoolean(value)

Case "compoundType"
    value = LCase(value)
    
    If value = "solvent" Or value = "reactant" Or value = "reagent" Then
        cmp.compoundType = value
    Else
        Dim inputValid As Boolean
        Dim inputVal As String
        Dim msg As String

        msg = "You entered an unrecognized compound type: '" & value & "'." & vbCrLf & vbCrLf & _
              "Please select a valid type by number:" & vbCrLf & _
              "[1] Reactant (for stoichiometric inputs)" & vbCrLf & _
              "[2] Reagent (for non-stoichiometric inputs)" & vbCrLf & _
              "[3] Solvent"

        inputValid = False

        Do While Not inputValid
            inputVal = InputBox(msg, "Invalid Compound Type")

            If inputVal = "" Then
                Debug.Print "[INFO] Compound type change skipped by user."
                Exit Do
            End If

            Select Case Trim(inputVal)
                Case "1"
                    cmp.compoundType = "reactant"
                    inputValid = True
                Case "2"
                    cmp.compoundType = "reagent"
                    inputValid = True
                Case "3"
                    cmp.compoundType = "solvent"
                    inputValid = True
                Case Else
                    MsgBox "Invalid choice. Please enter 1, 2, or 3.", vbExclamation, "Try Again"
            End Select
        Loop
    End If


            Case "corrected amount [g]"
                cmp.amount.SetCorrectedMass CDbl(value), cmp.Stoffdaten

            Case Else
                Debug.Print "[INFO] Unknown or unsupported key ignored: " & key
        End Select
SkipApply:
    Next i
End Sub

Private Function InferCompoundTypeFromCode(ByVal code As String) As String
    Dim prefix As String
    prefix = Left(code, 3)

    Select Case prefix
        Case "LM-", "SY-": InferCompoundTypeFromCode = "solvent"
        Case "SA-", "SL-": InferCompoundTypeFromCode = "reagent"
        Case "CH-", "LP-": InferCompoundTypeFromCode = "reactant"
        Case Else: InferCompoundTypeFromCode = ""
    End Select
End Function



'=== NEW: Refresh current Stoffdaten entry ===
Public Sub RefreshStoffdaten(ByVal cmp As Compound)
    Dim key As String
    Dim matches As Collection
    Dim selected As Stoffdaten
    Dim choice As Long

    ' Ensure we have a valid current product code
    If cmp.Stoffdaten Is Nothing Or cmp.Stoffdaten.productCode = "" Then
        MsgBox "Cannot refresh: Compound has no Stoffdaten or product code.", vbCritical, "RefreshStoffdaten"
        Err.Raise vbObjectError + 5001, "CompoundMutator.RefreshStoffdaten", _
                  "Missing Stoffdaten or product code"
    End If

    key = cmp.Stoffdaten.productCode
    Set matches = GetGlobalStoffdatenbank().FindMatchingStoffe(key)

    ' No matches ? abort with loud error
    If matches.count = 0 Then
        MsgBox "No Stoffdaten found for product code: " & key, vbCritical, "RefreshStoffdaten"
        Err.Raise vbObjectError + 5000, "CompoundMutator.RefreshStoffdaten", _
                  "No matching Stoffdaten for: " & key
    End If

    ' Single match ? apply and notify
    If matches.count = 1 Then
        Set selected = matches(1)
        ApplySelectedStoffdaten cmp, selected
        MsgBox "Stoffdaten for '" & key & "' refreshed.", vbInformation, "RefreshStoffdaten"
        Exit Sub
    End If

    ' Multiple matches ? prompt user selection
    choice = GetGlobalStoffdatenbank().PromptStoffIndexSelection(matches)
    If choice = 0 Then Exit Sub ' user cancelled
    Set selected = matches(choice)
    ApplySelectedStoffdaten cmp, selected
    MsgBox "Stoffdaten for '" & key & "' refreshed.", vbInformation, "RefreshStoffdaten"
End Sub


' Private helper to mirror UpdateProductCode post-lookup logic
Private Sub ApplySelectedStoffdaten(ByVal cmp As Compound, ByVal selected As Stoffdaten)
    ' Guard invalid molar mass
    If selected.molarMass = 0 Then
        MsgBox "Warning: Stoffdaten '" & selected.productCode & "' has molar mass = 0. Using -1.", vbExclamation, "RefreshStoffdaten"
        selected.molarMass = -1
    End If

    ' Guard invalid density
    If selected.Density = 0 Then
        MsgBox "Warning: Stoffdaten '" & selected.productCode & "' has density = 0. Using 1.", vbExclamation, "RefreshStoffdaten"
        selected.Density = 1
    End If

    ' Apply new Stoffdaten
    Set cmp.Stoffdaten = selected
' Only infer a new type if it isn’t already “product”
If cmp.compoundType <> "product" Then
    cmp.compoundType = InferCompoundTypeFromCode(selected.productCode)
End If
    cmp.InstructAssayCorrection = False

    ' Recalculate mass for new data
    cmp.amount.SetMass cmp.amount.mass, cmp.Stoffdaten
End Sub




