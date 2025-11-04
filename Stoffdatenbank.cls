VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Stoffdatenbank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' File: Stoffdatenbank.cls
Option Explicit

Private pStoffe As Collection
Private pIsInitialized As Boolean

Private Const DEFAULT_FILE_PATH As String = "https://cordenpharma.sharepoint.com/sites/cp_sui_intranetsui/prd/Freigegebene Dokumente/FE Stoffdatenbank.xlsm"
'Private Const DEFAULT_FILE_PATH As String = "C:\Users\max.klatte\Downloads\FE Stoffdatenbank (1).xlsm"
Private Const DEFAULT_SHEET_NAME As String = "Stoffdatenbank"

' === Public API ===

Public Function FindMatchingStoffe(ByVal searchTerm As String) As Collection
    Debug.Print "[DEBUG] Searching for: " & searchTerm
    Debug.Print "[DEBUG] Total loaded Stoffe: " & pStoffe.count

    Dim stoff As Stoffdaten
    Dim term As String
    Dim exactMatches As New Collection
    Dim partialMatches As New Collection
    Dim result As New Collection
    term = LCase(Trim(searchTerm))

    For Each stoff In pStoffe
        If LCase(stoff.productCode) = term _
        Or LCase(stoff.title) = term _
        Or LCase(stoff.CASReg) = term Then
            exactMatches.Add stoff
        ElseIf InStr(1, LCase(stoff.productCode), term) > 0 _
            Or InStr(1, LCase(stoff.title), term) > 0 _
            Or InStr(1, LCase(stoff.CASReg), term) > 0 Then
            partialMatches.Add stoff
        End If
    Next stoff

    For Each stoff In exactMatches: result.Add stoff: Next
    For Each stoff In partialMatches: result.Add stoff: Next

    Set FindMatchingStoffe = result
End Function

Public Function PromptStoffSelection() As Stoffdaten
    Dim searchTerm As String
    searchTerm = PromptSearchTerm()
    If searchTerm = "" Then Exit Function

    Dim resultList As Collection
    Set resultList = Me.FindMatchingStoffe(searchTerm)
    If resultList.count = 0 Then
        MsgBox "No matches found.", vbInformation
        Exit Function
    End If

    Dim choice As Variant
    choice = PromptStoffIndexSelection(resultList)
    If choice = 0 Then Exit Function

    Set PromptStoffSelection = resultList(choice)
End Function

Public Function LoadIfNotInitialized(Optional filePath As String = DEFAULT_FILE_PATH, _
                                     Optional sheetName As String = DEFAULT_SHEET_NAME, _
                                     Optional showConfirm As Boolean = True) As Boolean
    If pIsInitialized Then
        LoadIfNotInitialized = True
        Exit Function
    End If

    If showConfirm Then
        Dim userConfirm As VbMsgBoxResult
        userConfirm = MsgBox("This will load the Stoffdatenbank from SharePoint." & vbCrLf & _
                             "This may take a minute or more. Continue?", vbYesNo + vbExclamation, "Load Confirmation")
        If userConfirm <> vbYes Then
            Debug.Print "User cancelled Stoffdatenbank load."
            LoadIfNotInitialized = False
            Exit Function
        End If
    End If

    On Error GoTo LoadFail
    Dim startTime As Single: startTime = Timer
    Me.LoadStoffeFromExcel filePath, sheetName
    Dim elapsed As Single: elapsed = Timer - startTime

    MsgBox "Stoffdatenbank loaded in " & format(elapsed, "0.00") & " seconds.", vbInformation, "Load Complete"
    Debug.Print "[INFO] Loaded Stoffdatenbank in " & format(elapsed, "0.00") & " seconds."

    LoadIfNotInitialized = True
    Exit Function

LoadFail:
    MsgBox "Error loading Stoffdatenbank: " & Err.Description, vbCritical
    Debug.Print "[ERROR] Failed to load Stoffdatenbank: " & Err.Description
    LoadIfNotInitialized = False
End Function

Public Sub LoadStoffeFromExcel(ByVal filePath As String, ByVal sheetName As String)
    On Error GoTo LoadError
    Set pStoffe = New Collection

    Dim xlApp As Object, wb As Object, ws As Object
    Set wb = OpenWorkbook(filePath, xlApp)
    If wb Is Nothing Then Exit Sub

    Set ws = wb.Sheets(sheetName)
    Call ParseStoffeFromSheet(ws)

    pIsInitialized = True
    Debug.Print "[INFO] Loaded " & pStoffe.count & " rows."

Cleanup:
    SafeCloseExcel wb, xlApp
    Exit Sub

LoadError:
    Debug.Print "[FATAL] Load error: " & Err.Description
    Resume Cleanup
End Sub

Public Function GetGlobalStoffdatenbank() As Stoffdatenbank
    Static gStoffdatenbank As Stoffdatenbank
    If gStoffdatenbank Is Nothing Then Set gStoffdatenbank = New Stoffdatenbank

    If Not gStoffdatenbank.LoadIfNotInitialized Then
        Set gStoffdatenbank = Nothing
    End If
    Set GetGlobalStoffdatenbank = gStoffdatenbank
End Function

Public Property Get Stoffe() As Collection
    Set Stoffe = pStoffe
End Property

Public Property Get IsInitialized() As Boolean
    IsInitialized = pIsInitialized
End Property

' === Private Helpers ===

Private Function OpenWorkbook(ByVal filePath As String, ByRef xlApp As Object) As Object
    On Error GoTo HandleError
    Debug.Print "[INFO] Opening Excel: " & filePath
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    Set OpenWorkbook = xlApp.Workbooks.Open(filePath, ReadOnly:=True)
    Exit Function
HandleError:
    Debug.Print "[ERROR] Failed to open workbook: " & Err.Description
    Set OpenWorkbook = Nothing
End Function

Private Sub ParseStoffeFromSheet(ByVal ws As Object)
    On Error GoTo Fail

    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.count, "A").End(-4162).row
    If lastRow < 2 Then
        Debug.Print "[DEBUG] Sheet too short to parse. Last row: " & lastRow
        Exit Sub
    End If

    Dim data As Variant
    data = ws.Range("A2:F" & lastRow).value

    Debug.Print "[DEBUG] Row count loaded: " & UBound(data, 1)
    Debug.Print "[DEBUG] First row preview: " & Join(Array(data(1, 1), data(1, 2), data(1, 3)), " | ")

    Dim i As Long, stoff As Stoffdaten
    For i = 1 To UBound(data, 1)
        If Trim(data(i, 1)) <> "" Then
            Set stoff = BuildStoffFromRowData(data, i)
            pStoffe.Add stoff
        End If
    Next i
    Debug.Print "[DEBUG] Stoffe parsed: " & pStoffe.count
    Exit Sub

Fail:
    Debug.Print "[ERROR] Parse failed: " & Err.Description
End Sub

Private Function BuildStoffFromRowData(ByVal data As Variant, ByVal i As Long) As Stoffdaten
    Dim stoff As New Stoffdaten
    With stoff
        .productCode = data(i, 1)
        .title = SanitizeDisplayString(data(i, 2))
        .CASReg = data(i, 3)
        .Density = val(data(i, 5))
        .molarMass = val(data(i, 6))
    End With
    Set BuildStoffFromRowData = stoff
End Function

Private Sub SafeCloseExcel(ByRef wb As Object, ByRef xlApp As Object)
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    If Not xlApp Is Nothing Then xlApp.Quit
    Set wb = Nothing
    Set xlApp = Nothing
End Sub

Private Function PromptSearchTerm() As String
    PromptSearchTerm = Trim(InputBox("Enter a search term (Product Code, Title or CAS No):", "Search Stoffdaten"))
End Function

'=== Made Public: PromptStoffIndexSelection ===
Public Function PromptStoffIndexSelection(ByVal resultList As Collection) As Long
    Dim i As Long
    Dim listMsg As String
    listMsg = "Matches found: " & resultList.count & vbCrLf

    For i = 1 To resultList.count
        listMsg = listMsg & i & ". " & resultList(i).productCode & " | " & resultList(i).title & " | " & resultList(i).CASReg & vbCrLf
    Next i

    Dim choice As Variant
    choice = InputBox(listMsg & vbCrLf & "Please choose a number:", "Select Entry")
    If Not IsNumeric(choice) Or choice < 1 Or choice > resultList.count Then
        MsgBox "Invalid selection.", vbExclamation
        PromptStoffIndexSelection = 0
    Else
        PromptStoffIndexSelection = CLng(choice)
    End If
End Function


