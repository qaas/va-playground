Option Explicit

' =================================== GLOBAL VARIABLES ===================================
Dim gstrInputFileAPath As String
Dim gstrInputFileBPath As String
Dim gstrSupplementaryFileAPath As String
Dim gstrSupplementaryFileBPath As String
Dim gstrInputASource As String
Dim gstrInputBSource As String
Dim gstrInputADelimiter As String
Dim gstrInputBDelimiter As String
Dim gstrSupplementaryADelimiter As String
Dim gstrSupplementaryBDelimiter As String
Dim gstrPrimaryKeyA As String
Dim gstrPrimaryKeyB As String
Dim gstrSecondaryKeyA As String
Dim gstrSecondaryKeyB As String
Dim gstrSupplementaryKeyA As String
Dim gstrSupplementaryKeyB As String
Dim gstrSupplementaryASourceKey As String
Dim gstrSupplementaryBSourceKey As String

'temp function'
Sub UpdateMappingTableFormat()
    ' Updates the format of mapping tables in the Mapping worksheet
    On Error GoTo ErrorHandler
    
    Dim wsMappingDef As Worksheet
    Dim rangeToSearch As Range
    Dim cell As Range
    Dim i As Long
    Dim tableStartRows() As Long
    Dim tableStartCols() As Long
    Dim tableCount As Long
    Dim columnName As String
    
    ' Set worksheet
    Set wsMappingDef = Worksheets("Mapping")
    
    ' Find all mapping tables in the Mapping sheet
    tableCount = 0
    ReDim tableStartRows(1 To 100) ' Assume max 100 mapping tables
    ReDim tableStartCols(1 To 100) ' Store start column for each table
    
    ' Search for "Input File: " text to find mapping table starts
    Set rangeToSearch = wsMappingDef.UsedRange
    
    For Each cell In rangeToSearch.Cells
        If InStr(1, cell.value, "Input File:", vbTextCompare) > 0 Then
            tableCount = tableCount + 1
            tableStartRows(tableCount) = cell.Row
            tableStartCols(tableCount) = cell.Column
        End If
    Next cell
    
    ' If no mapping tables found, exit
    If tableCount = 0 Then
        MsgBox "No mapping tables found in the Mapping sheet.", vbInformation, "No Mapping Tables"
        Exit Sub
    End If
    
    ' Process each mapping table
    For i = 1 To tableCount
        Dim startRow As Long, startCol As Long
        startRow = tableStartRows(i)
        startCol = tableStartCols(i)
        
        ' Check if the table is already in the new format
        If wsMappingDef.Cells(startRow + 1, startCol).value = "Column Name in Source File" Then
            ' Current format - need to update:
            
            ' Get the column name from the current position (row+2)
            columnName = wsMappingDef.Cells(startRow + 2, startCol).value
            
            ' Store all mapping values temporarily
            Dim lastRow As Long
            lastRow = startRow + 2
            While wsMappingDef.Cells(lastRow + 1, startCol).value <> ""
                lastRow = lastRow + 1
            Wend
            
            ' Clear the headers and move everything up
            wsMappingDef.Cells(startRow + 1, startCol).value = columnName
            wsMappingDef.Cells(startRow + 1, startCol + 1).value = "Mapping Value"
            
            ' Move all mapping values up one row
            wsMappingDef.Range(wsMappingDef.Cells(startRow + 3, startCol), _
                              wsMappingDef.Cells(lastRow, startCol + 1)).Cut _
                              Destination:=wsMappingDef.Cells(startRow + 2, startCol)
            
            ' Clear the old last row that's now duplicated
            wsMappingDef.Cells(lastRow, startCol).ClearContents
            wsMappingDef.Cells(lastRow, startCol + 1).ClearContents
            
            ' Format the table
            wsMappingDef.Cells(startRow + 1, startCol).Font.Bold = True
            wsMappingDef.Cells(startRow + 1, startCol + 1).Font.Bold = True
            
            ' Apply borders to the entire table
            lastRow = startRow + 2
            While wsMappingDef.Cells(lastRow, startCol).value <> ""
                lastRow = lastRow + 1
            Wend
            lastRow = lastRow - 1
            
            wsMappingDef.Range(wsMappingDef.Cells(startRow, startCol), _
                              wsMappingDef.Cells(lastRow, startCol + 1)).Borders.LineStyle = xlContinuous
        End If
    Next i
    
    ' Format example tables in the Mapping worksheet
    FormatExampleMappingTables
    
    MsgBox "Mapping table format updated successfully!", vbInformation
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in UpdateMappingTableFormat", vbCritical, "Error"
End Sub


Sub UpdateApplyMappingFunction()
    ' Update code in the ApplyMapping function to handle the new table format
    MsgBox "Please add the following changes to your ApplyMapping function:" & vbCrLf & vbCrLf & _
           "1. Change 'columnName = wsMappingDef.Cells(tableStartRows(i) + 2, startCol).Value' to:" & vbCrLf & _
           "   columnName = wsMappingDef.Cells(tableStartRows(i) + 1, startCol).Value" & vbCrLf & vbCrLf & _
           "2. Change 'j = tableStartRow + 3' to:" & vbCrLf & _
           "   j = tableStartRow + 2" & vbCrLf & vbCrLf & _
           "This will adjust the code to work with the new table format.", vbInformation, _
           "ApplyMapping Function Updates Needed"
End Sub

Sub UpdateMappingWorksheet()
    ' Run both updates
    UpdateMappingTableFormat
    UpdateApplyMappingFunction
End Sub
' =================================== INITIALIZATION ===================================
Sub InitializeWorkbook()
    ' This macro creates all the required worksheets and formats them according to specifications
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Create all required worksheets in a safer way
    Call SafeCreateAllWorksheets
    
    ' Format worksheets
    Call FormatWorksheets
    
    ' Create buttons on Parameters sheet
    Call CreateButtons
    
CleanExit:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    MsgBox "Workbook initialized successfully!", vbInformation, "Initialization Complete"
    Exit Sub
    
ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & vbNewLine & _
           "Try closing and reopening Excel, then run the macro again.", vbCritical, "Initialization Error"
    Resume CleanExit
End Sub

Function WorksheetExists(wsName As Variant) As Boolean
    ' Check if worksheet exists
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = Worksheets(CStr(wsName))
    On Error GoTo 0
    
    WorksheetExists = Not ws Is Nothing
End Function

Sub SafeCreateAllWorksheets()
    ' Create all required worksheets in a safer way
    Dim wsNames As Variant
    Dim wsColors As Variant
    Dim i As Integer
    Dim wsExists As Boolean
    Dim newWs As Worksheet
    
    
    ' *** Change the order of the worksheet names and colours if worksheet position change is required. *** '
    
    ' Define worksheet names
    wsNames = Array("Instructions", "Input Parameters", "Mapping", "Result", "Dashboard", "Input File A", "Input File B", _
                   "Supplementary File A", "Supplementary File B")
    
    ' Define tab colors (RGB format)
    wsColors = Array(RGB(255, 255, 153), RGB(204, 255, 204), RGB(192, 192, 192), RGB(255, 204, 153), RGB(255, 255, 204), _
                     RGB(204, 204, 255), RGB(255, 153, 153), RGB(204, 255, 255), _
                     RGB(255, 204, 255))
    
    ' Check if we have any of the required worksheets already
    For i = 0 To UBound(wsNames)
        If WorksheetExists(wsNames(i)) Then
            ' If the sheet exists, clear its contents but keep the sheet
            On Error Resume Next
            Worksheets(wsNames(i)).Cells.Clear
            ' Try to set the tab color if possible
            Worksheets(wsNames(i)).Tab.Color = wsColors(i)
            On Error GoTo 0
        Else
            ' Add the sheet if it doesn't exist
            On Error Resume Next
            Set newWs = Worksheets.Add(After:=Worksheets(Worksheets.count))
            If Err.Number = 0 Then
                newWs.Name = wsNames(i)
                ' Try to set the tab color if possible
                newWs.Tab.Color = wsColors(i)
            End If
            On Error GoTo 0
        End If
    Next i
    
    ' Handle the case where Sheet1 still exists and isn't one of our named sheets
    If WorksheetExists("Sheet1") Then
        Dim sheetFound As Boolean
        sheetFound = False
        
        For i = 0 To UBound(wsNames)
            If wsNames(i) = "Sheet1" Then
                sheetFound = True
                Exit For
            End If
        Next i
        
        If Not sheetFound Then
            ' If Sheet1 isn't one of our named sheets, rename it to the first sheet we need
            On Error Resume Next
            Worksheets("Sheet1").Name = wsNames(0)
            Worksheets(wsNames(0)).Tab.Color = wsColors(0)
            On Error GoTo 0
        End If
    End If
End Sub

Sub FormatWorksheets()
    ' Format all worksheets according to specifications
    On Error Resume Next
    
    ' 1. Instructions worksheet
    If WorksheetExists("Instructions") Then
        Call FormatInstructionsSheet
    End If
    
    ' 2. Input Parameters worksheet
    If WorksheetExists("Input Parameters") Then
        Call FormatParametersSheet
    End If
    
    ' 3. Result worksheet
    If WorksheetExists("Result") Then
        Call FormatResultSheet
    End If
    
    ' 4-7. Input and Supplementary files worksheets (minimal formatting, content will be populated later)
    If WorksheetExists("Input File A") Then
        Call FormatDataSheet("Input File A")
    End If
    
    If WorksheetExists("Input File B") Then
        Call FormatDataSheet("Input File B")
    End If
    
    If WorksheetExists("Supplementary File A") Then
        Call FormatDataSheet("Supplementary File A")
    End If
    
    If WorksheetExists("Supplementary File B") Then
        Call FormatDataSheet("Supplementary File B")
    End If
    
    ' 8. Mapping worksheet
    If WorksheetExists("Mapping") Then
        Call FormatMappingSheet
    End If
    
    ' 9. Dashboard worksheet
    If WorksheetExists("Dashboard") Then
        Call FormatDashboardSheet
    End If
    
    On Error GoTo 0
End Sub

Sub FormatInstructionsSheet()
    ' Format Instructions worksheet
    Dim ws As Worksheet
    Set ws = Worksheets("Instructions")
    
    ' Clear sheet
    ws.Cells.Clear
    
    ' Add title
    ws.Range("A1").value = "RECONCILIATION TOOL - INSTRUCTIONS"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 16
    
    ' Add instructions
    ws.Range("A3").value = "HOW TO USE THIS RECONCILIATION TOOL:"
    ws.Range("A3").Font.Bold = True
    
    ws.Range("A5").value = "1. Go to 'Input Parameters' tab and fill in the required information:"
    ws.Range("A6").value = "   - Input File Paths"
    ws.Range("A7").value = "   - Data Source Names"
    ws.Range("A8").value = "   - File Delimiters"
    ws.Range("A9").value = "   - Reconciliation Keys"
    ws.Range("A10").value = "   - Mapping Information (if needed)"
    
    ws.Range("A12").value = "2. Use the buttons on the 'Input Parameters' tab to run each step of the reconciliation process:"
    ws.Range("A13").value = "   - Load Input Files: Imports data from the specified input files"
    ws.Range("A14").value = "   - Load Supplementary Files: Imports data from supplementary files (if specified)"
    ws.Range("A15").value = "   - Apply Mapping: Applies any mapping defined in the 'Mapping' tab"
    ws.Range("A16").value = "   - Run Reconciliation: Performs the reconciliation based on specified keys"
    ws.Range("A17").value = "   - Generate Dashboard: Creates a summary of reconciliation results"
    
    ws.Range("A19").value = "3. View the results in the 'Result' tab and 'Dashboard' tab"
    
    ws.Range("A21").value = "NOTES:"
    ws.Range("A21").Font.Bold = True
    ws.Range("A22").value = "- All data is processed as text to preserve leading zeros"
    ws.Range("A23").value = "- For better performance with large files, run each step separately"
    ws.Range("A24").value = "- Make sure column names in mapping tables match exactly with input files"
    
    ' Auto-fit columns
    ws.Columns("A:A").AutoFit
End Sub

Sub FormatParametersSheet()
    ' Format Input Parameters worksheet
    Dim ws As Worksheet
    Set ws = Worksheets("Input Parameters")
    Dim i As Integer
    
    ' Clear sheet
    ws.Cells.Clear
    
    ' Add title
    ws.Range("A1").value = "RECONCILIATION TOOL - INPUT PARAMETERS"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 16
    ws.Range("A1:F1").Merge
    
    ' Format sections
    
    ' 1. Input File Paths
    ws.Range("A3").value = "1. INPUT FILE PATHS"
    ws.Range("A3").Font.Bold = True
    
    ws.Range("A5").value = "Input File A"
    ws.Range("B5").value = ""
    ws.Range("C5").Interior.Color = RGB(204, 204, 255)  ' Match 'Input File A' tab color
    
    ws.Range("A6").value = "Input File B"
    ws.Range("B6").value = ""
    ws.Range("C6").Interior.Color = RGB(255, 153, 153)  ' Match 'Input File B' tab color
    
    ws.Range("A7").value = "Supplementary File A (Optional)"
    ws.Range("B7").value = ""
    ws.Range("C7").Interior.Color = RGB(204, 255, 255)  ' Match 'Supplementary File A' tab color
    
    ws.Range("A8").value = "Supplementary File B (Optional)"
    ws.Range("B8").value = ""
    ws.Range("C8").Interior.Color = RGB(255, 204, 255)  ' Match 'Supplementary File B' tab color
    
    ' 2. Input Names
    ws.Range("A10").value = "2. DATA SOURCE NAMES"
    ws.Range("A10").Font.Bold = True
    
    ws.Range("A12").value = "Source of Input File A"
    ws.Range("B12").value = ""
    ws.Range("C12").Interior.Color = RGB(204, 204, 255)
    
    ws.Range("A13").value = "Source of Input File B"
    ws.Range("B13").value = ""
    ws.Range("C13").Interior.Color = RGB(255, 153, 153)
    
    ' 3. Delimiters
    ws.Range("A15").value = "3. FILE DELIMITERS"
    ws.Range("A15").Font.Bold = True
    
    ws.Range("A17").value = "Delimiter of Input File A"
    ws.Range("B17").value = ","
    
    ws.Range("A18").value = "Delimiter of Input File B"
    ws.Range("B18").value = ","
    
    ws.Range("A19").value = "Delimiter of Supplementary File A"
    ws.Range("B19").value = ","
    
    ws.Range("A20").value = "Delimiter of Supplementary File B"
    ws.Range("B20").value = ","
    
    ' 4. Reconciliation Keys
    ws.Range("A22").value = "4. RECONCILIATION KEYS"
    ws.Range("A22").Font.Bold = True
    
    ws.Range("A24").value = "Column Name for Recon Keys"
    ws.Range("B24").value = "Primary Key"
    ws.Range("C24").value = "Secondary Key"
    ws.Range("A24:C24").Font.Bold = True
    
    ws.Range("A25").value = "Input File A"
    ws.Range("B25").value = ""
    ws.Range("C25").value = ""
    
    ws.Range("A26").value = "Input File B"
    ws.Range("B26").value = ""
    ws.Range("C26").value = ""
    
    ' 5. Supplementary Data Keys
    ws.Range("A28").value = "5. SUPPLEMENTARY DATA A KEYS"
    ws.Range("A28").Font.Bold = True
    
    ws.Range("A30").value = "Column Name for Supplementary Data A Keys"
    ws.Range("B30").value = "Primary Key"
    ws.Range("A30:B30").Font.Bold = True
    
    ws.Range("A31").value = "Input File A"
    ws.Range("B31").value = ""
    
    ws.Range("A32").value = "Supplementary File A"
    ws.Range("B32").value = ""
    
    ' 6. Supplementary Data Columns
    ws.Range("A34").value = "Supplementary Data A Column Required at Output"
    ws.Range("A34").Font.Bold = True
    
    ' Add a few blank rows for user input
    ws.Range("A35").value = ""
    ws.Range("A36").value = ""
    ws.Range("A37").value = ""
    
    ' 7. Supplementary Data B Keys
    ws.Range("A39").value = "6. SUPPLEMENTARY DATA B KEYS"
    ws.Range("A39").Font.Bold = True
    
    ws.Range("A41").value = "Column Name for Supplementary Data B Keys"
    ws.Range("B41").value = "Primary Key"
    ws.Range("A41:B41").Font.Bold = True
    
    ws.Range("A42").value = "Input File B"
    ws.Range("B42").value = ""
    
    ws.Range("A43").value = "Supplementary File B"
    ws.Range("B43").value = ""
    
    ' 8. Supplementary Data Columns
    ws.Range("A45").value = "Supplementary Data B Column Required at Output"
    ws.Range("A45").Font.Bold = True
    
    ' Add a few blank rows for user input
    ws.Range("A46").value = ""
    ws.Range("A47").value = ""
    ws.Range("A48").value = ""
    
    ' Add Column Mapping Table in columns I-J
    ws.Range("I3").value = "COLUMN MAPPING TABLE"
    ws.Range("I3").Font.Bold = True
    
    ws.Range("I5").value = "Input File A"
    ws.Range("J5").value = "Input File B"
    ws.Range("I5:J5").Font.Bold = True
    
    ' Add empty rows for mapping (can be pre-populated by users)
    For i = 6 To 25
        ws.Range("I" & i).value = ""
        ws.Range("J" & i).value = ""
    Next i
    
    ' Format the mapping table
    ws.Range("I5:J25").Borders.LineStyle = xlContinuous
    
    ' Format all cells with borders
    ws.Range("A3:C48").Borders.LineStyle = xlContinuous
    
    ' Auto-fit columns
    ws.Columns("A:J").AutoFit
    
    'Add red asterisks to mandatory fields
    AddMandatoryFieldIndicators
    
End Sub

Sub AddMandatoryFieldIndicators()
    ' Add red asterisks to mandatory fields in the Input Parameters sheet
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim r As Range
    Dim mandatoryFields As Variant
    Dim i As Long
    
    ' List of mandatory fields to mark with asterisks - field labels in column A
    mandatoryFields = Array("Input File A", "Input File B", _
                           "Source of Input File A", "Source of Input File B", _
                           "Delimiter of Input File A", "Delimiter of Input File B", _
                           "Column Name for Recon Keys")
    
    ' Set worksheet
    Set ws = Worksheets("Input Parameters")
    
    ' Loop through each mandatory field
    For i = LBound(mandatoryFields) To UBound(mandatoryFields)
        ' Search for the label in column A
        Set r = ws.Columns("A").Find(What:=mandatoryFields(i), _
                                    LookIn:=xlValues, _
                                    LookAt:=xlWhole, _
                                    SearchOrder:=xlByRows, _
                                    SearchDirection:=xlNext, _
                                    MatchCase:=False)
        
        ' If found, add a red asterisk
        If Not r Is Nothing Then
            ' Check if it already has an asterisk
            If Right(r.value, 1) <> "*" Then
                ' Add asterisk
                r.value = r.value & " *"
                
                ' Format the asterisk in red
                With r.Characters(Start:=Len(r.value), Length:=1).Font
                    .Color = RGB(255, 0, 0)
                    .Bold = True
                End With
            End If
        End If
    Next i
    
    ' Handle the Primary Key field specifically - it's in cell B24
    If Not IsEmpty(ws.Range("B24").value) Then
        If Right(ws.Range("B24").value, 1) <> "*" Then
            ws.Range("B24").value = ws.Range("B24").value & " *"
            
            With ws.Range("B24").Characters(Start:=Len(ws.Range("B24").value), Length:=1).Font
                .Color = RGB(255, 0, 0)
                .Bold = True
            End With
        End If
    Else
        ws.Range("B24").value = "Primary Key *"
        
        With ws.Range("B24").Characters(Start:=Len(ws.Range("B24").value), Length:=1).Font
            .Color = RGB(255, 0, 0)
            .Bold = True
        End With
    End If
    
    ' Also handle the column mapping table in columns I & J
    ws.Range("I5").value = "Input File A *"
    ws.Range("J5").value = "Input File B *"
    
    ' Format the asterisks in red
    With ws.Range("I5").Characters(Start:=Len(ws.Range("I5").value), Length:=1).Font
        .Color = RGB(255, 0, 0)
        .Bold = True
    End With
    
    With ws.Range("J5").Characters(Start:=Len(ws.Range("J5").value), Length:=1).Font
        .Color = RGB(255, 0, 0)
        .Bold = True
    End With
    
    ' Add a note about mandatory fields
    ws.Range("A1").Offset(0, 1).value = "* Mandatory fields"
    ws.Range("A1").Offset(0, 1).Font.Color = RGB(255, 0, 0)
    ws.Range("A1").Offset(0, 1).Font.Bold = True
        
    Exit Sub
    
ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in AddMandatoryFieldIndicators", vbCritical, "Error"
End Sub


Sub FormatDataSheet(sheetName As String)
    ' Basic formatting for data sheets
    Dim ws As Worksheet
    Set ws = Worksheets(sheetName)
    
    ' Clear sheet
    ws.Cells.Clear
    
    ' Add title
    ws.Range("A1").value = "DATA FROM " & UCase(sheetName)
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
    
    ' Add note
    ws.Range("A2").value = "This data will be populated when you load the input files."
    ws.Range("A2").Font.Italic = True
End Sub

Sub FormatMappingSheet()
    ' Format Mapping worksheet
    Dim ws As Worksheet
    Set ws = Worksheets("Mapping")
    
    ' Clear sheet
    ws.Cells.Clear
    
    ' Add title
    ws.Range("A1").value = "MAPPING TABLES"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 16
    
    ' Add instructions
    ws.Range("A3").value = "Instructions:"
    ws.Range("A3").Font.Bold = True
    ws.Range("A4").value = "- Add mapping tables below for any columns that require value mapping before reconciliation"
    ws.Range("A5").value = "- Each mapping table must include the input file (A or B) and the column name"
    ws.Range("A6").value = "- Separate mapping tables with at least one empty column"
    
    ' Example mapping table
    ws.Range("A8").value = "Input File: <A/B>"
    ws.Range("A8").Font.Bold = True
    
    ws.Range("A9").value = "<Column Name in Source File>"
    ws.Range("B9").value = "Mapping Value"
    ws.Range("A9:B9").Font.Bold = True
    
    ws.Range("A10").value = "<Source Value 1>"
    ws.Range("B10").value = "<Mapped Value 1>"
    
    ws.Range("A11").value = "<Source Value 2>"
    ws.Range("B11").value = "<Mapped Value 2>"
    
    ' Add another example
    ws.Range("D8").value = "Input File: <A/B>"
    ws.Range("D8").Font.Bold = True
    
    ws.Range("D9").value = "<Column Name in Source File>"
    ws.Range("E9").value = "Mapping Value"
    ws.Range("D9:E9").Font.Bold = True
    
    ws.Range("D10").value = "<Source Value 1>"
    ws.Range("E10").value = "<Mapped Value 1>"
    
    ws.Range("D11").value = "<Source Value 2>"
    ws.Range("E11").value = "<Mapped Value 2>"
    
    ' Format table with borders
    ws.Range("A8:B11").Borders.LineStyle = xlContinuous
    ws.Range("D8:E11").Borders.LineStyle = xlContinuous
    
    ' Auto-fit columns
    ws.Columns("A:E").AutoFit
End Sub

Sub FormatResultSheet()
    ' Basic formatting for Result sheet
    Dim ws As Worksheet
    Set ws = Worksheets("Result")
    
    ' Clear sheet
    ws.Cells.Clear
    
    ' Add title
    ws.Range("A1").value = "RECONCILIATION RESULTS"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 16
    
    ' Add note
    ws.Range("A2").value = "Results will be populated when you run the reconciliation process."
    ws.Range("A2").Font.Italic = True
End Sub

Sub FormatDashboardSheet()
    ' Basic formatting for Dashboard sheet
    Dim ws As Worksheet
    Set ws = Worksheets("Dashboard")
    
    ' Clear sheet
    ws.Cells.Clear
    
    ' Add title
    ws.Range("A1").value = "RECONCILIATION DASHBOARD"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 16
    
    ' Add note
    ws.Range("A2").value = "Dashboard will be populated after reconciliation is complete."
    ws.Range("A2").Font.Italic = True
End Sub

Sub CreateButtons()
    ' Create buttons on the Parameters sheet
    On Error Resume Next
    
    Dim ws As Worksheet
    If Not WorksheetExists("Input Parameters") Then
        MsgBox "Input Parameters sheet doesn't exist. Unable to create buttons.", vbExclamation
        Exit Sub
    End If
    
    Set ws = Worksheets("Input Parameters")
    
    ' Add a section for buttons
    ws.Range("E3").value = "RUN RECONCILIATION STEPS"
    ws.Range("E3").Font.Bold = True
    
    ' Remove any existing buttons to avoid duplicates
    ClearExistingButtons ws
    
    ' Create buttons for each step
    ' Note: In macOS, ActiveX controls might not work properly, so we use Form controls
    
    ' Button 1: Load Input Files
    AddButton ws, 340, 30, 120, 30, "ThisWorkbook.LoadInputFiles", "1. Load Input Files", "btnLoadInputFiles"
    
    ' Button 2: Load Supplementary Files
    AddButton ws, 340, 65, 120, 30, "ThisWorkbook.LoadSupplementaryFiles", "2. Load Supplementary Files", "btnLoadSupplementaryFiles"
    
    ' Button 3: Apply Mapping
    AddButton ws, 340, 100, 120, 30, "ThisWorkbook.ApplyMapping", "3. Apply Mapping", "btnApplyMapping"
    
    ' Button 4: Run Reconciliation
    AddButton ws, 340, 135, 120, 30, "ThisWorkbook.RunReconciliation", "4. Run Reconciliation", "btnRunReconciliation"
    
    ' Button 5: Generate Dashboard
    AddButton ws, 340, 170, 120, 30, "ThisWorkbook.GenerateDashboard", "5. Generate Dashboard", "btnGenerateDashboard"
    
    ' Button 6: Run All Steps
    AddButton ws, 340, 215, 120, 30, "ThisWorkbook.RunAllSteps", "Run All Steps", "btnRunAllSteps"
    
    ' Deselect buttons
    ws.Range("A1").Select
        
    On Error GoTo 0
End Sub

Sub ClearExistingButtons(ws As Worksheet)
    ' Remove any existing buttons from the worksheet
    On Error Resume Next
    
    ' Clear form buttons
    Dim btn As Button
    For Each btn In ws.Buttons
        btn.Delete
    Next btn
    
    On Error GoTo 0
End Sub

Sub AddButton(ws As Worksheet, left As Double, top As Double, width As Double, height As Double, _
              actionName As String, captionText As String, buttonName As String)
    ' Add a button safely with error handling
    On Error Resume Next
    
    Dim btn As Button
    Set btn = ws.Buttons.Add(left, top, width, height)
    
    If Not btn Is Nothing Then
        btn.OnAction = actionName
        btn.Caption = captionText
        btn.Name = buttonName
    End If
    
    On Error GoTo 0
End Sub

' =================================== LOAD FILE FUNCTIONS ===================================

Sub LoadInputFiles()
    ' Load input files A and B
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    
    ' Get parameters from Input Parameters sheet
    ReadParameters
    
    ' Clear any mapped columns from previous runs
    ClearMappedColumns Worksheets("Input File A")
    ClearMappedColumns Worksheets("Input File B")
    
    ' Load File A
    If gstrInputFileAPath <> "" Then
        LoadDataFromFile gstrInputFileAPath, "Input File A", gstrInputADelimiter
    Else
        MsgBox "Input File A path is missing. Please provide a valid file path.", vbExclamation, "Missing Input"
        Exit Sub
    End If
    
    ' Load File B
    If gstrInputFileBPath <> "" Then
        LoadDataFromFile gstrInputFileBPath, "Input File B", gstrInputBDelimiter
    Else
        MsgBox "Input File B path is missing. Please provide a valid file path.", vbExclamation, "Missing Input"
        Exit Sub
    End If
    
    ' Create mapping table in Mapping sheet for reconciliation columns
    CreateMappingTable
    
CleanExit:
    Application.ScreenUpdating = True
    
    If Err.Number = 0 Then
        MsgBox "Input files loaded successfully!", vbInformation, "Files Loaded"
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in LoadInputFiles", vbCritical, "Error"
    Resume CleanExit
End Sub

Sub LoadSupplementaryFiles()
    ' Load supplementary files if specified
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    
    ' Get parameters from Input Parameters sheet
    ReadParameters
    
    ' Clear any mapped columns from previous runs
    ClearMappedColumns Worksheets("Supplementary File A")
    ClearMappedColumns Worksheets("Supplementary File B")
    
    ' Load Supplementary File A if path is provided
    If gstrSupplementaryFileAPath <> "" Then
        LoadDataFromFile gstrSupplementaryFileAPath, "Supplementary File A", gstrSupplementaryADelimiter
        MsgBox "Supplementary File A loaded successfully!", vbInformation, "File Loaded"
    Else
        MsgBox "Supplementary File A path is not provided. Skipping this step.", vbInformation, "Optional Step Skipped"
    End If
    
    ' Load Supplementary File B if path is provided
    If gstrSupplementaryFileBPath <> "" Then
        LoadDataFromFile gstrSupplementaryFileBPath, "Supplementary File B", gstrSupplementaryBDelimiter
        MsgBox "Supplementary File B loaded successfully!", vbInformation, "File Loaded"
    Else
        MsgBox "Supplementary File B path is not provided. Skipping this step.", vbInformation, "Optional Step Skipped"
    End If
    
CleanExit:
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in LoadSupplementaryFiles", vbCritical, "Error"
    Resume CleanExit
End Sub

Sub ReadParameters()
    ' Read parameters from Input Parameters sheet
    Dim ws As Worksheet
    Set ws = Worksheets("Input Parameters")
    
    ' Read file paths
    gstrInputFileAPath = ws.Range("B5").value
    gstrInputFileBPath = ws.Range("B6").value
    gstrSupplementaryFileAPath = ws.Range("B7").value
    gstrSupplementaryFileBPath = ws.Range("B8").value
    
    ' Read source names
    gstrInputASource = ws.Range("B12").value
    gstrInputBSource = ws.Range("B13").value
    
    ' Read delimiters
    gstrInputADelimiter = ws.Range("B17").value
    gstrInputBDelimiter = ws.Range("B18").value
    gstrSupplementaryADelimiter = ws.Range("B19").value
    gstrSupplementaryBDelimiter = ws.Range("B20").value
    
    ' Read reconciliation keys
    gstrPrimaryKeyA = ws.Range("B25").value
    gstrPrimaryKeyB = ws.Range("B26").value
    gstrSecondaryKeyA = ws.Range("C25").value
    gstrSecondaryKeyB = ws.Range("C26").value
    
    ' Read supplementary keys
    gstrSupplementaryKeyA = ws.Range("B31").value
    gstrSupplementaryASourceKey = ws.Range("B32").value
    gstrSupplementaryKeyB = ws.Range("B42").value
    gstrSupplementaryBSourceKey = ws.Range("B43").value
End Sub

Function IsDate2(strValue As String) As Boolean
    ' More conservative date detection function
    On Error Resume Next
    
    IsDate2 = False
    
    ' Skip empty strings
    If Trim(strValue) = "" Then Exit Function
    
    ' If it's purely numeric with no separators, it's probably not a date
    If IsNumeric(strValue) Then
        ' Pure integers are not treated as dates
        If InStr(1, strValue, ".") = 0 And InStr(1, strValue, "/") = 0 And InStr(1, strValue, "-") = 0 Then
            IsDate2 = False
            Exit Function
        End If
    End If
    
    ' Check for date separators - must have at least one to be considered a date
    If InStr(1, strValue, "/") = 0 And InStr(1, strValue, "-") = 0 Then
        IsDate2 = False
        Exit Function
    End If
    
    ' Check if it's a standard date format
    If IsDate(strValue) Then
        ' Additional validation to avoid false positives
        ' Must contain digits and separators in patterns typical of dates
        If (InStr(1, strValue, "/") > 0 Or InStr(1, strValue, "-") > 0) And _
           Len(strValue) >= 6 Then ' At least M/D/YY format
            IsDate2 = True
        End If
        Exit Function
    End If
    
    ' Handle potential ISO date format (YYYY-MM-DD)
    If Len(strValue) = 10 Then
        If Mid(strValue, 5, 1) = "-" And Mid(strValue, 8, 1) = "-" Then
            Dim year As String, month As String, day As String
            year = left(strValue, 4)
            month = Mid(strValue, 6, 2)
            day = Right(strValue, 2)
            
            If IsNumeric(year) And IsNumeric(month) And IsNumeric(day) Then
                If CInt(year) >= 1900 And CInt(year) <= 2100 Then
                    If CInt(month) >= 1 And CInt(month) <= 12 Then
                        If CInt(day) >= 1 And CInt(day) <= 31 Then
                            IsDate2 = True
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    On Error GoTo 0
End Function

Sub LoadDataFromFile(filePath As String, destSheet As String, delimiter As String)
    ' Load data from a CSV/TXT file to a worksheet
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim fileNum As Integer
    Dim dataLine As String
    Dim dataArray() As String
    Dim i As Long, j As Long
    Dim rowCount As Long
    Dim cellValue As String
    Dim isDateValue As Boolean
    Dim dateHeaders As New Collection
    
    ' Set destination worksheet
    Set ws = Worksheets(destSheet)
    
    ' Clear previous data (both content and formatting)
    ws.Range("A3:ZZ" & ws.Rows.count).ClearContents
    ws.Range("A3:ZZ" & ws.Rows.count).ClearFormats
    
    ' Format header row again
    ws.Rows(3).Font.Bold = True
    
    ' Get a free file number
    fileNum = FreeFile
    
    ' Open the file for input
    On Error Resume Next
    Open filePath For Input As #fileNum
    
    If Err.Number <> 0 Then
        MsgBox "Error opening file: " & filePath & vbCrLf & _
               "Error: " & Err.Description, vbCritical, "File Error"
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    
    ' Start at row 3 (row 1-2 are for title and instructions)
    rowCount = 3
    
    ' Read the file line by line
    Do Until EOF(fileNum)
        Line Input #fileNum, dataLine
        
        ' Skip empty lines
        If Trim(dataLine) <> "" Then
            ' Split the line using the specified delimiter
            dataArray = Split(dataLine, delimiter)
            
            ' Write each element to a cell
            For j = 0 To UBound(dataArray)
                ' Trim the data value
                cellValue = Trim(dataArray(j))
                
                ' Set the cell value
                ws.Cells(rowCount, j + 1).value = cellValue
                
                ' First row (headers) gets text format
                If rowCount = 3 Then
                    ws.Cells(rowCount, j + 1).NumberFormat = "@"
                Else
                    ' Check if it looks like a date
                    isDateValue = IsDate2(cellValue)
                    
                    ' Apply appropriate formatting
                    If isDateValue Then
                        ' Format as date if it appears to be a date
                        On Error Resume Next
                        ws.Cells(rowCount, j + 1).NumberFormat = "yyyy-mm-dd"
                        
                        ' Remember this column as a date column
                        On Error Resume Next
                        dateHeaders.Add j + 1, "Col" & (j + 1)
                        On Error GoTo ErrorHandler
                    Else
                        ' Otherwise, treat as text (to preserve leading zeros)
                        ws.Cells(rowCount, j + 1).NumberFormat = "@"
                    End If
                End If
            Next j
            
            ' Move to next row
            rowCount = rowCount + 1
        End If
    Loop
    
    ' Close the file
    Close #fileNum
    
    ' Format the header row
    ws.Range("A3:" & ws.Cells(3, j).Address).Font.Bold = True
    ws.Range("A3:" & ws.Cells(3, j).Address).Interior.Color = RGB(220, 220, 220)
    
    ' Auto-fit columns
    ws.Columns("A:Z").AutoFit
    
    Exit Sub
    
ErrorHandler:
    ' Close the file if it's open
    On Error Resume Next
    Close #fileNum
    On Error GoTo 0
    
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in LoadDataFromFile", vbCritical, "Error"
End Sub

Sub CopyValueWithFormatting(sourceWs As Worksheet, sourceRow As Long, sourceCol As Long, destWs As Worksheet, destRow As Long, destCol As Long)
    ' Copy a cell value while preserving formatting (especially date formats)
    On Error Resume Next
    
    ' Copy the value
    destWs.Cells(destRow, destCol).value = sourceWs.Cells(sourceRow, sourceCol).value
    
    ' Preserve the number format
    destWs.Cells(destRow, destCol).NumberFormat = sourceWs.Cells(sourceRow, sourceCol).NumberFormat
    
    On Error GoTo 0
End Sub

Sub ClearMappedColumns(ws As Worksheet)
    ' Clears any mapped columns (those ending with _map) from the worksheet
    On Error GoTo ErrorHandler
    
    Dim lastCol As Long
    Dim i As Long
    Dim j As Long
    Dim colsToRemove As New Collection
    
    ' Find last column with data
    lastCol = ws.Cells(3, ws.Columns.count).End(xlToLeft).Column
    
    ' First, collect indexes of mapped columns
    For i = 1 To lastCol
        Dim header As String
        header = ws.Cells(3, i).value
        
        If Len(header) >= 4 Then
            If Right(header, 4) = "_map" Then
                ' Add to collection of columns to remove
                On Error Resume Next
                colsToRemove.Add i
                On Error GoTo ErrorHandler
            End If
        End If
    Next i
    
    ' Now clear those columns (content and formatting)
    If colsToRemove.count > 0 Then
        ' Process in reverse order to avoid shifting issues
        For i = colsToRemove.count To 1 Step -1
            Dim colIndex As Long
            colIndex = colsToRemove(i)
            
            ' Clear the column
            ws.Columns(colIndex).ClearContents
            ws.Columns(colIndex).ClearFormats
            
            ' Delete the column (optional - only if you want to remove them completely)
            ' ws.Columns(colIndex).Delete shift:=xlToLeft
        Next i
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in ClearMappedColumns", vbCritical, "Error"
End Sub

Sub CreateMappingTable()
    ' Create the reconciliation column mapping table in columns I and J of the Input Parameters worksheet
    Dim wsMapping As Worksheet
    Dim wsFileA As Worksheet
    Dim lastColA As Long
    Dim i As Long
    Dim rowNum As Long
    Dim isMappingTableEmpty As Boolean
    
    ' Set worksheets
    Set wsMapping = Worksheets("Input Parameters")
    Set wsFileA = Worksheets("Input File A")
    
    ' Check if the mapping table is already populated
    isMappingTableEmpty = True
    For i = 6 To 25  ' Check rows 6-25
        If wsMapping.Range("I" & i).value <> "" Then
            isMappingTableEmpty = False
            Exit For
        End If
    Next i
    
    ' Only populate the table if it's empty
    If isMappingTableEmpty Then
        ' Find the last column with data in row 3 (header row)
        lastColA = wsFileA.Cells(3, wsFileA.Columns.count).End(xlToLeft).Column
        
        ' Start adding column mappings from row 6
        rowNum = 6
        
        ' Add File A columns
        For i = 1 To lastColA
            wsMapping.Cells(rowNum, 9).value = wsFileA.Cells(3, i).value  ' Column I is index 9
            rowNum = rowNum + 1
            ' Don't exceed the prepared rows (6-25)
            If rowNum > 25 Then Exit For
        Next i
    End If
    
    ' Format the range with borders
    wsMapping.Range("I5:J25").Borders.LineStyle = xlContinuous
    
    ' Auto-fit columns
    wsMapping.Columns("I:J").AutoFit
End Sub

' =================================== MAPPING FUNCTIONS ===================================

Sub ApplyMapping()
    ' Apply mapping defined in the Mapping sheet
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Dim wsMappingDef As Worksheet
    Dim wsMapping As Worksheet
    Dim rangeToSearch As Range
    Dim cell As Range
    Dim i As Long, j As Long
    Dim lastRow As Long, lastCol As Long
    Dim tableStartRows() As Long
    Dim tableStartCols() As Long  ' Store the column where each table starts
    Dim tableCount As Long
    Dim fileType As String
    Dim columnName As String
    Dim sourceValue As String
    Dim mappedValue As String
    Dim wsInput As Worksheet
    
    ' Set worksheets
    On Error Resume Next
    Set wsMappingDef = Worksheets("Mapping")
    Set wsMapping = Worksheets("Input Parameters")
    On Error GoTo ErrorHandler
    
    If wsMappingDef Is Nothing Then
        MsgBox "Worksheet 'Mapping' not found.", vbExclamation, "Missing Worksheet"
        GoTo CleanExit
    End If
    
    If wsMapping Is Nothing Then
        MsgBox "Worksheet 'Input Parameters' not found.", vbExclamation, "Missing Worksheet"
        GoTo CleanExit
    End If
    
    ' Check if the column mapping table has been filled out
    Dim hasMappings As Boolean
    
    hasMappings = False
    
    ' Check for at least one mapping pair
    For i = 6 To 25  ' Check rows 6-25 (mapping table rows)
        If wsMapping.Cells(i, 9).value <> "" And wsMapping.Cells(i, 10).value <> "" Then
            hasMappings = True
            Exit For
        End If
    Next i
    
    If Not hasMappings Then
        MsgBox "The column mapping table is not filled out. Please go to the Input Parameters tab " & _
               "and map columns from Input File A to Input File B in columns I and J.", _
               vbExclamation, "Missing Mappings"
        GoTo CleanExit
    End If
    
    ' Find all mapping tables in the Mapping sheet
    tableCount = 0
    ReDim tableStartRows(1 To 100) ' Assume max 100 mapping tables
    ReDim tableStartCols(1 To 100) ' Store start column for each table
    
    ' Search for "Input File: " text to find mapping table starts
    Set rangeToSearch = wsMappingDef.UsedRange
    
    For Each cell In rangeToSearch.Cells
        If InStr(1, cell.value, "Input File:", vbTextCompare) > 0 Then
            tableCount = tableCount + 1
            tableStartRows(tableCount) = cell.Row
            tableStartCols(tableCount) = cell.Column ' Store the column
        End If
    Next cell
    
    ' If no mapping tables found, exit
    If tableCount = 0 Then
        MsgBox "No mapping tables found in the Mapping sheet. Skipping this step.", vbInformation, "No Mapping"
        GoTo CleanExit
    End If
    
    ' Process each mapping table
    For i = 1 To tableCount
        ' Get file type (A or B) using safer method
        Dim headerText As String
        Dim startCol As Long
        
        startCol = tableStartCols(i) ' Use the stored column
        headerText = Trim(wsMappingDef.Cells(tableStartRows(i), startCol).value)
        
        fileType = ""
        If InStr(1, headerText, "Input File: A", vbTextCompare) > 0 Or _
           InStr(1, headerText, "Input File:A", vbTextCompare) > 0 Then
            fileType = "A"
        ElseIf InStr(1, headerText, "Input File: B", vbTextCompare) > 0 Or _
               InStr(1, headerText, "Input File:B", vbTextCompare) > 0 Then
            fileType = "B"
        End If
        
        ' Validation
        If fileType = "" Then
            MsgBox "Invalid file type in mapping table at row " & tableStartRows(i) & ", column " & startCol & vbCrLf & _
                   "Cell contains: """ & headerText & """" & vbCrLf & _
                   "Please use exactly 'Input File: A' or 'Input File: B'", _
                   vbExclamation, "Mapping Error"
            GoTo CleanExit
        End If
        
        ' UPDATED FOR NEW FORMAT: Get column name from row +1 instead of +2
        ' Get column name from the header row (one row below file type)
        columnName = wsMappingDef.Cells(tableStartRows(i) + 1, startCol).value
        
        ' Debugging message
        Debug.Print "Processing mapping table at row " & tableStartRows(i) & ", column " & startCol & _
                    ", File Type: " & fileType & ", Column Name: " & columnName
        
        ' Column name validation
        If columnName = "" Then
            MsgBox "Missing column name in mapping table at row " & (tableStartRows(i) + 1) & ", column " & startCol & vbCrLf & _
                   "Please specify which column to map in the row below the file type.", _
                   vbExclamation, "Mapping Error"
            GoTo CleanExit
        End If
        
        ' Set the input worksheet based on file type
        If fileType = "A" Then
            Set wsInput = Worksheets("Input File A")
        ElseIf fileType = "B" Then
            Set wsInput = Worksheets("Input File B")
        End If
        
        If wsInput Is Nothing Then
            MsgBox "Worksheet 'Input File " & fileType & "' not found.", vbExclamation, "Missing Worksheet"
            GoTo CleanExit
        End If
        
        ' Find the column with the specified name in the input worksheet
        lastCol = wsInput.Cells(3, wsInput.Columns.count).End(xlToLeft).Column
        Dim columnFound As Boolean
        columnFound = False
        
        For j = 1 To lastCol
            If wsInput.Cells(3, j).value = columnName Then
                ' Found the column, create a mapped column
                AddMappedColumn wsInput, j, tableStartRows(i), wsMappingDef, startCol
                columnFound = True
                Exit For
            End If
        Next j
        
        ' Column not found warning
        If Not columnFound Then
            MsgBox "Column '" & columnName & "' not found in " & IIf(fileType = "A", "Input File A", "Input File B") & _
                   ". Please check the column name.", vbExclamation, "Column Not Found"
        End If
    Next i
    
CleanExit:
    ' Always restore Excel settings
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    If Err.Number = 0 Then
        MsgBox "Mapping applied successfully!", vbInformation, "Mapping Complete"
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Error in Apply Mapping"
    Resume CleanExit
End Sub

Sub AddMappedColumn(wsInput As Worksheet, columnIndex As Long, tableStartRow As Long, wsMappingDef As Worksheet, startCol As Long)
    ' Create a new column with mapped values instead of replacing original values
    On Error GoTo ErrorHandler
    
    Dim lastRowInput As Long
    Dim i As Long, j As Long
    Dim sourceValue As String
    Dim mappedValue As String
    Dim mappingFound As Boolean
    Dim lastRowMapping As Long
    Dim columnName As String
    Dim lastCol As Long
    Dim newColIndex As Long
    Dim mappedColName As String
    
    ' Get the original column name
    columnName = wsInput.Cells(3, columnIndex).value
    
    ' Create the new column name with _map suffix
    mappedColName = columnName & "_map"
    
    ' Find the last column with data
    lastCol = wsInput.Cells(3, wsInput.Columns.count).End(xlToLeft).Column
    
    ' Check if the mapped column already exists
    Dim mappedColExists As Boolean
    mappedColExists = False
    
    For j = 1 To lastCol
        If wsInput.Cells(3, j).value = mappedColName Then
            newColIndex = j
            mappedColExists = True
            Exit For
        End If
    Next j
    
    ' If mapped column doesn't exist, add it at the end
    If Not mappedColExists Then
        newColIndex = lastCol + 1
        
        ' Add column header with _map suffix
        wsInput.Cells(3, newColIndex).value = mappedColName
        
        ' Format the header with a darker color to distinguish it
        wsInput.Cells(3, newColIndex).Interior.Color = RGB(153, 153, 204) ' Darker shade
        wsInput.Cells(3, newColIndex).Font.Bold = True
    End If
    
    ' Find the last row with data in the input worksheet
    lastRowInput = wsInput.Cells(wsInput.Rows.count, columnIndex).End(xlUp).Row
    
    ' Find the last row of the mapping table - starting from data row after the column name header
    ' UPDATED FOR NEW FORMAT: Start from row +2 instead of +3
    j = tableStartRow + 2 ' First mapping value row (in new format)
    
    Do While j <= wsMappingDef.Rows.count And wsMappingDef.Cells(j, startCol).value <> ""
        j = j + 1
    Loop
    lastRowMapping = j - 1
    
    ' Copy original values to mapped column and apply mapping
    For i = 4 To lastRowInput ' Start from row 4 (skip header in row 3)
        sourceValue = wsInput.Cells(i, columnIndex).value
        
        ' Skip empty cells
        If sourceValue <> "" Then
            ' Look for the source value in the mapping table
            mappingFound = False
            
            ' UPDATED FOR NEW FORMAT: Start from row +2 instead of +3
            For j = tableStartRow + 2 To lastRowMapping ' Start from first mapping value row
                If wsMappingDef.Cells(j, startCol).value = sourceValue Then
                    ' Found a mapping
                    mappedValue = wsMappingDef.Cells(j, startCol + 1).value
                    wsInput.Cells(i, newColIndex).value = mappedValue
                    mappingFound = True
                    Exit For
                End If
            Next j
            
            ' If no mapping found, mark the cell
            If Not mappingFound Then
                wsInput.Cells(i, newColIndex).value = "N/A - failed to map. Original Value: " & sourceValue
                wsInput.Cells(i, newColIndex).Interior.Color = RGB(255, 191, 0) ' Amber color
            End If
        Else
            ' For empty cells, leave the mapped column empty
            wsInput.Cells(i, newColIndex).value = ""
        End If
    Next i
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & _
           "While mapping column " & columnIndex & " in " & wsInput.Name, vbCritical, "Error in Apply Mapping"
End Sub



' =================================== RECONCILIATION FUNCTIONS ===================================

Sub RunReconciliation()
    ' Run the reconciliation process
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    
    ' Declare all variables
    Dim ws As Worksheet
    Dim i As Long
    Dim hasMappings As Boolean
    Dim mappings() As String
    
    ' Read parameters
    ReadParameters
    
    ' Check if primary keys are specified
    If gstrPrimaryKeyA = "" Or gstrPrimaryKeyB = "" Then
        MsgBox "Primary keys for reconciliation are not specified. Please provide them in the Input Parameters sheet.", _
               vbExclamation, "Missing Keys"
        GoTo CleanExit
    End If
    
    ' Check if the column mapping table has been filled out
    Set ws = Worksheets("Input Parameters")
    hasMappings = False
    
    ' Check for at least one mapping pair
    For i = 6 To 25  ' Check rows 6-25 (mapping table rows)
        If ws.Cells(i, 9).value <> "" And ws.Cells(i, 10).value <> "" Then
            hasMappings = True
            Exit For
        End If
    Next i
    
    If Not hasMappings Then
        MsgBox "The column mapping table is not filled out. Please go to the Input Parameters tab " & _
               "and map columns from Input File A to Input File B in columns I and J.", _
               vbExclamation, "Missing Mappings"
        GoTo CleanExit
    End If
    
    ' Get column mappings from Input Parameters sheet
    mappings = GetColumnMappings()
    
    ' Check if we got any mappings (array size)
    If UBound(mappings, 1) <= 0 And UBound(mappings, 2) <= 0 Then
        GoTo CleanExit  ' GetColumnMappings already showed an error message
    End If
    
    ' Perform reconciliation with the mappings
    PerformReconciliation mappings
    
CleanExit:
    Application.ScreenUpdating = True
    
    If Err.Number = 0 Then
        MsgBox "Reconciliation completed successfully!", vbInformation, "Reconciliation Complete"
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Error in Reconciliation"
    Resume CleanExit
End Sub

'This function clears Result tab before populating it with new reconciliation results.
Sub ClearResultSheet()
    ' Clear both content and formatting from the Result sheet
    On Error GoTo ErrorHandler
    
    Dim wsResult As Worksheet
    
    ' Set worksheet
    Set wsResult = Worksheets("Result")
    
    ' Preserve title and subtitle
    Dim title As String, subtitle As String
    title = wsResult.Range("A1").value
    subtitle = wsResult.Range("A2").value
    
    ' Clear everything from row 3 down (both content and formatting)
    wsResult.Range("A3:ZZ" & wsResult.Rows.count).ClearContents
    wsResult.Range("A3:ZZ" & wsResult.Rows.count).ClearFormats
    
    ' Restore title and subtitle
    wsResult.Range("A1").value = title
    wsResult.Range("A2").value = subtitle
    wsResult.Range("A1").Font.Bold = True
    wsResult.Range("A1").Font.Size = 16
    wsResult.Range("A2").Font.Italic = True
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in ClearResultSheet", vbCritical, "Error"
End Sub

' This function gets the mappings as two parallel arrays in a single 2D array
Function GetColumnMappings() As String()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim i As Long, count As Long
    Dim mappings() As String
    
    ' Set worksheet
    Set ws = Worksheets("Input Parameters")
    
    ' First count how many mappings we have
    count = 0
    For i = 6 To 25
        If ws.Cells(i, 9).value <> "" And ws.Cells(i, 10).value <> "" Then
            count = count + 1
        End If
    Next i
    
    If count = 0 Then
        MsgBox "The column mapping table is not filled out. Please go to the Input Parameters tab " & _
               "and map columns from Input File A to Input File B in columns I and J.", _
               vbExclamation, "Missing Mappings"
        ReDim mappings(0 To 0, 0 To 1)  ' Return empty array with proper dimensions
        GetColumnMappings = mappings
        Exit Function
    End If
    
    ' Create a 2D array to hold mappings - first column for A, second for B
    ReDim mappings(0 To count - 1, 0 To 1)
    
    ' Read mappings into the 2D array
    count = 0
    For i = 6 To 25
        If ws.Cells(i, 9).value <> "" And ws.Cells(i, 10).value <> "" Then
            mappings(count, 0) = ws.Cells(i, 9).value
            mappings(count, 1) = ws.Cells(i, 10).value
            count = count + 1
        End If
    Next i
    
    GetColumnMappings = mappings
    Exit Function
    
ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in GetColumnMappings", vbCritical, "Error"
    ReDim mappings(0 To 0, 0 To 1)  ' Return empty array with proper dimensions
    GetColumnMappings = mappings
End Function

Sub PerformReconciliation(mappings() As String)
    ' Perform the reconciliation and populate the Result sheet using a 2D array of mappings
    On Error GoTo ErrorHandler
    
    ' All variable declarations remain the same
    Dim wsFileA As Worksheet
    Dim wsFileB As Worksheet
    Dim wsResult As Worksheet
    Dim wsSupplA As Worksheet
    Dim wsSupplB As Worksheet
    Dim lastRowA As Long, lastRowB As Long
    Dim lastColA As Long, lastColB As Long
    Dim lastColSupplA As Long, lastColSupplB As Long
    Dim i As Long, j As Long, k As Long
    Dim rowResult As Long
    Dim colPrimaryKeyA As Long, colPrimaryKeyB As Long
    Dim colSecondaryKeyA As Long, colSecondaryKeyB As Long
    Dim colSupplKeyA As Long, colSupplKeyB As Long
    Dim primaryKeyValueA As String, primaryKeyValueB As String
    Dim secondaryKeyValueA As String, secondaryKeyValueB As String
    Dim matchFound As Boolean
    Dim matchType As String
    Dim matchRowB As Long
    Dim colA As Long, colB As Long
    Dim valueA As String, valueB As String
    Dim dataA() As Variant
    Dim dataB() As Variant
    Dim supplDataA() As Variant
    Dim supplDataB() As Variant
    Dim headerA() As String
    Dim headerB() As String
    Dim headerSupplA() As String
    Dim headerSupplB() As String
    Dim requiredSupplColsA() As String
    Dim requiredSupplColsB() As String
    Dim requiredSupplColIndexA() As Long
    Dim requiredSupplColIndexB() As Long
    Dim hasSupplA As Boolean
    Dim hasSupplB As Boolean
    Dim lastRowSupplA As Long
    Dim lastRowSupplB As Long
    Dim supplMatchRowA As Long
    Dim supplMatchRowB As Long
    Dim singletonSupplMatchRowB As Long
    Dim reconColCount As Long
    Dim colOffset As Long
    Dim columnFound As Boolean
    Dim mappingCount As Long
    
    ' New variables for tracking mismatched fields
    Dim mismatchedFields As String
    Dim hasMismatch As Boolean
    
    ' New variables for mapped columns
    Dim colAComp As Long, colBComp As Long
    Dim mappedColNameA As String, mappedColNameB As String
    Dim hasMappedCol As Boolean
    Dim mappedColIndex As Long
    
    ' Safely get mapping count
    On Error Resume Next
    mappingCount = UBound(mappings, 1) + 1
    If Err.Number <> 0 Then
        mappingCount = 0
    End If
    On Error GoTo ErrorHandler
    
    ' Set worksheets with error handling
    On Error Resume Next
    Set wsFileA = Worksheets("Input File A")
    Set wsFileB = Worksheets("Input File B")
    Set wsResult = Worksheets("Result")
    Set wsSupplA = Worksheets("Supplementary File A")
    Set wsSupplB = Worksheets("Supplementary File B")
    
    If wsFileA Is Nothing Then
        MsgBox "Worksheet 'Input File A' not found.", vbExclamation, "Missing Worksheet"
        Exit Sub
    End If
    
    If wsFileB Is Nothing Then
        MsgBox "Worksheet 'Input File B' not found.", vbExclamation, "Missing Worksheet"
        Exit Sub
    End If
    
    If wsResult Is Nothing Then
        MsgBox "Worksheet 'Result' not found.", vbExclamation, "Missing Worksheet"
        Exit Sub
    End If
    
    On Error GoTo ErrorHandler
    
    ' Clear Result sheet (both content and formatting)
    ClearResultSheet
    
    ' Find the last rows and columns with data
    lastRowA = wsFileA.Cells(wsFileA.Rows.count, 1).End(xlUp).Row
    lastRowB = wsFileB.Cells(wsFileB.Rows.count, 1).End(xlUp).Row
    lastColA = wsFileA.Cells(3, wsFileA.Columns.count).End(xlToLeft).Column
    lastColB = wsFileB.Cells(3, wsFileB.Columns.count).End(xlToLeft).Column
    
    ' Check if data exists
    If lastRowA <= 3 Then
        MsgBox "No data found in Input File A.", vbExclamation, "No Data"
        Exit Sub
    End If
    
    If lastRowB <= 3 Then
        MsgBox "No data found in Input File B.", vbExclamation, "No Data"
        Exit Sub
    End If
    
    ' Find the primary and secondary key columns
    colPrimaryKeyA = FindColumnIndex(wsFileA, gstrPrimaryKeyA)
    colPrimaryKeyB = FindColumnIndex(wsFileB, gstrPrimaryKeyB)
    
    If gstrSecondaryKeyA <> "" Then
        colSecondaryKeyA = FindColumnIndex(wsFileA, gstrSecondaryKeyA)
    Else
        colSecondaryKeyA = 0
    End If
    
    If gstrSecondaryKeyB <> "" Then
        colSecondaryKeyB = FindColumnIndex(wsFileB, gstrSecondaryKeyB)
    Else
        colSecondaryKeyB = 0
    End If
    
    ' Validate key columns
    If colPrimaryKeyA = 0 Then
        MsgBox "Primary key column '" & gstrPrimaryKeyA & "' not found in Input File A.", vbExclamation, "Key Not Found"
        Exit Sub
    End If
    
    If colPrimaryKeyB = 0 Then
        MsgBox "Primary key column '" & gstrPrimaryKeyB & "' not found in Input File B.", vbExclamation, "Key Not Found"
        Exit Sub
    End If
    
    If gstrSecondaryKeyA <> "" And colSecondaryKeyA = 0 Then
        MsgBox "Secondary key column '" & gstrSecondaryKeyA & "' not found in Input File A.", vbExclamation, "Key Not Found"
        Exit Sub
    End If
    
    If gstrSecondaryKeyB <> "" And colSecondaryKeyB = 0 Then
        MsgBox "Secondary key column '" & gstrSecondaryKeyB & "' not found in Input File B.", vbExclamation, "Key Not Found"
        Exit Sub
    End If
    
    ' Check supplementary files
    hasSupplA = False
    hasSupplB = False
    
    ' Safely check if supplementary files exist
    On Error Resume Next
    hasSupplA = (gstrSupplementaryFileAPath <> "" And Not wsSupplA Is Nothing)
    hasSupplB = (gstrSupplementaryFileBPath <> "" And Not wsSupplB Is Nothing)
    On Error GoTo ErrorHandler
    
    ' Find supplementary key columns if needed
    If hasSupplA Then
        On Error Resume Next
        lastColSupplA = wsSupplA.Cells(3, wsSupplA.Columns.count).End(xlToLeft).Column
        colSupplKeyA = FindColumnIndex(wsSupplA, gstrSupplementaryASourceKey)
        
        If Err.Number <> 0 Or colSupplKeyA = 0 Then
            If Err.Number = 0 Then
                MsgBox "Supplementary key column '" & gstrSupplementaryASourceKey & "' not found in Supplementary File A.", _
                      vbExclamation, "Key Not Found"
            Else
                MsgBox "Error accessing Supplementary File A: " & Err.Description, vbExclamation, "Error"
            End If
            hasSupplA = False
        End If
        
        ' Get required supplementary columns for File A
        On Error Resume Next
        requiredSupplColsA = GetRequiredSupplementaryColumns("A")
        If Err.Number <> 0 Then
            MsgBox "Error getting supplementary columns for File A: " & Err.Description, vbExclamation, "Error"
            ReDim requiredSupplColsA(0)
            requiredSupplColsA(0) = ""
        End If
        
        On Error GoTo ErrorHandler
        
        ' Check if requiredSupplColsA is properly initialized
        Dim suppColsAFound As Boolean
        suppColsAFound = False
        
        On Error Resume Next
        If UBound(requiredSupplColsA) >= 0 Then
            suppColsAFound = True
        End If
        On Error GoTo ErrorHandler
        
        If hasSupplA And suppColsAFound Then
            On Error Resume Next
            ReDim requiredSupplColIndexA(UBound(requiredSupplColsA))
            
            For i = 0 To UBound(requiredSupplColsA)
                requiredSupplColIndexA(i) = FindColumnIndex(wsSupplA, requiredSupplColsA(i))
                
                If requiredSupplColIndexA(i) = 0 Then
                    ' Skip reporting error for empty column names
                    If requiredSupplColsA(i) <> "" Then
                        MsgBox "Required supplementary column '" & requiredSupplColsA(i) & "' not found in Supplementary File A.", _
                               vbExclamation, "Column Not Found"
                    End If
                End If
            Next i
            On Error GoTo ErrorHandler
        End If
    End If
    
    If hasSupplB Then
        On Error Resume Next
        lastColSupplB = wsSupplB.Cells(3, wsSupplB.Columns.count).End(xlToLeft).Column
        colSupplKeyB = FindColumnIndex(wsSupplB, gstrSupplementaryBSourceKey)
        
        If Err.Number <> 0 Or colSupplKeyB = 0 Then
            If Err.Number = 0 Then
                MsgBox "Supplementary key column '" & gstrSupplementaryBSourceKey & "' not found in Supplementary File B.", _
                      vbExclamation, "Key Not Found"
            Else
                MsgBox "Error accessing Supplementary File B: " & Err.Description, vbExclamation, "Error"
            End If
            hasSupplB = False
        End If
        
        ' Get required supplementary columns for File B
        On Error Resume Next
        requiredSupplColsB = GetRequiredSupplementaryColumns("B")
        If Err.Number <> 0 Then
            MsgBox "Error getting supplementary columns for File B: " & Err.Description, vbExclamation, "Error"
            ReDim requiredSupplColsB(0)
            requiredSupplColsB(0) = ""
        End If
        
        On Error GoTo ErrorHandler
        
        ' Check if requiredSupplColsB is properly initialized
        Dim suppColsBFound As Boolean
        suppColsBFound = False
        
        On Error Resume Next
        If UBound(requiredSupplColsB) >= 0 Then
            suppColsBFound = True
        End If
        On Error GoTo ErrorHandler
        
        If hasSupplB And suppColsBFound Then
            On Error Resume Next
            ReDim requiredSupplColIndexB(UBound(requiredSupplColsB))
            
            For i = 0 To UBound(requiredSupplColsB)
                requiredSupplColIndexB(i) = FindColumnIndex(wsSupplB, requiredSupplColsB(i))
                
                If requiredSupplColIndexB(i) = 0 Then
                    ' Skip reporting error for empty column names
                    If requiredSupplColsB(i) <> "" Then
                        MsgBox "Required supplementary column '" & requiredSupplColsB(i) & "' not found in Supplementary File B.", _
                               vbExclamation, "Column Not Found"
                    End If
                End If
            Next i
            On Error GoTo ErrorHandler
        End If
    End If
    
    ' Get headers from both files
    ReDim headerA(1 To lastColA)
    ReDim headerB(1 To lastColB)
    
    For i = 1 To lastColA
        headerA(i) = wsFileA.Cells(3, i).value
    Next i
    
    For i = 1 To lastColB
        headerB(i) = wsFileB.Cells(3, i).value
    Next i
    
    ' Get headers from supplementary files if needed
    If hasSupplA Then
        On Error Resume Next
        ReDim headerSupplA(1 To lastColSupplA)
        
        For i = 1 To lastColSupplA
            headerSupplA(i) = wsSupplA.Cells(3, i).value
        Next i
        On Error GoTo ErrorHandler
    End If
    
    If hasSupplB Then
        On Error Resume Next
        ReDim headerSupplB(1 To lastColSupplB)
        
        For i = 1 To lastColSupplB
            headerSupplB(i) = wsSupplB.Cells(3, i).value
        Next i
        On Error GoTo ErrorHandler
    End If
    
    ' Load data into arrays for faster processing
    ReDim dataA(4 To lastRowA, 1 To lastColA)
    ReDim dataB(4 To lastRowB, 1 To lastColB)
    
    For i = 4 To lastRowA
        For j = 1 To lastColA
            dataA(i, j) = Trim(wsFileA.Cells(i, j).value)
        Next j
    Next i
    
    For i = 4 To lastRowB
        For j = 1 To lastColB
            dataB(i, j) = Trim(wsFileB.Cells(i, j).value)
        Next j
    Next i
    
    ' Load supplementary data if needed
    If hasSupplA Then
        On Error Resume Next
        lastRowSupplA = wsSupplA.Cells(wsSupplA.Rows.count, 1).End(xlUp).Row
        
        If lastRowSupplA > 3 Then
            ReDim supplDataA(4 To lastRowSupplA, 1 To lastColSupplA)
            
            For i = 4 To lastRowSupplA
                For j = 1 To lastColSupplA
                    supplDataA(i, j) = Trim(wsSupplA.Cells(i, j).value)
                Next j
            Next i
        Else
            hasSupplA = False
        End If
        On Error GoTo ErrorHandler
    End If
    
    If hasSupplB Then
        On Error Resume Next
        lastRowSupplB = wsSupplB.Cells(wsSupplB.Rows.count, 1).End(xlUp).Row
        
        If lastRowSupplB > 3 Then
            ReDim supplDataB(4 To lastRowSupplB, 1 To lastColSupplB)
            
            For i = 4 To lastRowSupplB
                For j = 1 To lastColSupplB
                    supplDataB(i, j) = Trim(wsSupplB.Cells(i, j).value)
                Next j
            Next i
        Else
            hasSupplB = False
        End If
        On Error GoTo ErrorHandler
    End If
    
    ' Set up Result sheet headers
    ' Column 1: Reconciliation Result
    wsResult.Cells(3, 1).value = "Recon Result"
    wsResult.Columns("A:A").ColumnWidth = 50  ' Increase the width for detailed mismatch info
    
    ' Columns for File A data - include both original and mapped columns
    colOffset = 2 ' Start after Recon Result column
    
    For i = 1 To lastColA
        ' Check if this is a mapped column (ends with _map)
        If Right(wsFileA.Cells(3, i).value, 4) <> "_map" Then
            wsResult.Cells(3, colOffset).value = wsFileA.Cells(3, i).value & " (" & gstrInputASource & ")"
            wsResult.Cells(3, colOffset).Interior.Color = RGB(204, 204, 255) ' Match Input File A color
            colOffset = colOffset + 1
            
            ' Check if there's a mapped version of this column
            mappedColNameA = wsFileA.Cells(3, i).value & "_map"
            For j = 1 To lastColA
                If wsFileA.Cells(3, j).value = mappedColNameA Then
                    ' Add the mapped column too
                    wsResult.Cells(3, colOffset).value = mappedColNameA & " (" & gstrInputASource & ")"
                    wsResult.Cells(3, colOffset).Interior.Color = RGB(153, 153, 204) ' Darker shade for mapped columns
                    wsResult.Cells(3, colOffset).Font.Bold = True
                    colOffset = colOffset + 1
                    Exit For
                End If
            Next j
        End If
    Next i
    
    ' Columns for Supplementary File A data if needed
    Dim suppColsAInitialized As Boolean
    suppColsAInitialized = False
    
    On Error Resume Next
    If hasSupplA And UBound(requiredSupplColsA) >= 0 Then
        suppColsAInitialized = True
    End If
    On Error GoTo ErrorHandler
    
    If hasSupplA And suppColsAInitialized Then
        For i = 0 To UBound(requiredSupplColsA)
            If requiredSupplColsA(i) <> "" Then
                wsResult.Cells(3, colOffset).value = requiredSupplColsA(i) & " (Suppl " & gstrInputASource & ")"
                wsResult.Cells(3, colOffset).Interior.Color = RGB(204, 255, 255) ' Match Supplementary File A color
                colOffset = colOffset + 1
            End If
        Next i
    End If
    
    ' Columns for File B data - include both original and mapped columns
    For i = 1 To lastColB
        ' Check if this is a mapped column (ends with _map)
        If Right(wsFileB.Cells(3, i).value, 4) <> "_map" Then
            wsResult.Cells(3, colOffset).value = wsFileB.Cells(3, i).value & " (" & gstrInputBSource & ")"
            wsResult.Cells(3, colOffset).Interior.Color = RGB(255, 153, 153) ' Match Input File B color
            colOffset = colOffset + 1
            
            ' Check if there's a mapped version of this column
            mappedColNameB = wsFileB.Cells(3, i).value & "_map"
            For j = 1 To lastColB
                If wsFileB.Cells(3, j).value = mappedColNameB Then
                    ' Add the mapped column too
                    wsResult.Cells(3, colOffset).value = mappedColNameB & " (" & gstrInputBSource & ")"
                    wsResult.Cells(3, colOffset).Interior.Color = RGB(153, 153, 204) ' Darker shade for mapped columns
                    wsResult.Cells(3, colOffset).Font.Bold = True
                    colOffset = colOffset + 1
                    Exit For
                End If
            Next j
        End If
    Next i
    
    ' Columns for Supplementary File B data if needed
    Dim suppColsBInitialized As Boolean
    suppColsBInitialized = False
    
    On Error Resume Next
    If hasSupplB And UBound(requiredSupplColsB) >= 0 Then
        suppColsBInitialized = True
    End If
    On Error GoTo ErrorHandler
    
    If hasSupplB And suppColsBInitialized Then
        For i = 0 To UBound(requiredSupplColsB)
            If requiredSupplColsB(i) <> "" Then
                wsResult.Cells(3, colOffset).value = requiredSupplColsB(i) & " (Suppl " & gstrInputBSource & ")"
                wsResult.Cells(3, colOffset).Interior.Color = RGB(255, 204, 255) ' Match Supplementary File B color
                colOffset = colOffset + 1
            End If
        Next i
    End If
    
    ' Columns for reconciliation results
    reconColCount = 0
    
    ' Loop through the mappings array
    If mappingCount > 0 Then
        For i = 0 To UBound(mappings, 1)
            colA = FindColumnIndex(wsFileA, mappings(i, 0))
            colB = FindColumnIndex(wsFileB, mappings(i, 1))
            
            ' Try to find mapped columns
            mappedColNameA = mappings(i, 0) & "_map"
            mappedColNameB = mappings(i, 1) & "_map"
            
            colAComp = FindColumnIndex(wsFileA, mappedColNameA)
            colBComp = FindColumnIndex(wsFileB, mappedColNameB)
            
            ' If either original column exists, add a recon column
            If colA > 0 And colB > 0 Then
                wsResult.Cells(3, colOffset).value = mappings(i, 0) & " vs " & mappings(i, 1) & " (Recon)"
                wsResult.Cells(3, colOffset).Interior.Color = RGB(255, 204, 153) ' Match Result color
                colOffset = colOffset + 1
                reconColCount = reconColCount + 1
            End If
        Next i
    End If
    
    ' Start populating reconciliation results
    rowResult = 4
    
    ' Process each record in File A
    For i = 4 To lastRowA
        ' Get key values
        primaryKeyValueA = dataA(i, colPrimaryKeyA)
        
        If colSecondaryKeyA > 0 Then
            secondaryKeyValueA = dataA(i, colSecondaryKeyA)
        Else
            secondaryKeyValueA = ""
        End If
        
        ' Try to find a match in File B
        matchFound = False
        matchType = ""
        matchRowB = 0
        
        ' First try to match with both primary and secondary keys if available
        If colSecondaryKeyA > 0 And colSecondaryKeyB > 0 Then
            For j = 4 To lastRowB
                If dataB(j, colPrimaryKeyB) = primaryKeyValueA And _
                   dataB(j, colSecondaryKeyB) = secondaryKeyValueA Then
                    matchFound = True
                    matchType = "Full Match"
                    matchRowB = j
                    Exit For
                End If
            Next j
        End If
        
        ' If no match found with both keys, try primary key only
        If Not matchFound Then
            For j = 4 To lastRowB
                If dataB(j, colPrimaryKeyB) = primaryKeyValueA Then
                    matchFound = True
                    matchType = "Primary Key Match"
                    matchRowB = j
                    Exit For
                End If
            Next j
        End If
        
        ' Write the result to the Results sheet
        ' Recon Result column
        If matchFound Then
            ' Pre-check for mismatches to enhance the result message
            mismatchedFields = ""
            hasMismatch = False
            
            ' Check each mapped column for mismatches
            If mappingCount > 0 Then
                For k = 0 To UBound(mappings, 1)
                    colA = FindColumnIndex(wsFileA, mappings(k, 0))
                    colB = FindColumnIndex(wsFileB, mappings(k, 1))
                    
                    If colA > 0 And colB > 0 Then
                        ' Try to find mapped columns for comparison
                        mappedColNameA = mappings(k, 0) & "_map"
                        mappedColNameB = mappings(k, 1) & "_map"
                        
                        colAComp = FindColumnIndex(wsFileA, mappedColNameA)
                        colBComp = FindColumnIndex(wsFileB, mappedColNameB)
                        
                        ' Use mapped columns if available, otherwise use original
                        If colAComp > 0 Then
                            valueA = wsFileA.Cells(i, colAComp).value
                        Else
                            valueA = wsFileA.Cells(i, colA).value
                        End If
                        
                        If colBComp > 0 Then
                            valueB = wsFileB.Cells(matchRowB, colBComp).value
                        Else
                            valueB = wsFileB.Cells(matchRowB, colB).value
                        End If
                        
                        If valueA <> valueB Then
                            If hasMismatch Then
                                mismatchedFields = mismatchedFields & ", "
                            End If
                            mismatchedFields = mismatchedFields & mappings(k, 0)
                            hasMismatch = True
                        End If
                    End If
                Next k
            End If
            
            ' Enhanced result message with field mismatches
            If hasMismatch Then
                wsResult.Cells(rowResult, 1).value = matchType & " - Mismatched Fields: " & mismatchedFields
            Else
                wsResult.Cells(rowResult, 1).value = matchType & " - All Fields Match"
            End If
            
            ' Write File A data to Result sheet - both original and mapped columns
            colOffset = 2 ' Start after Recon Result column
            
            ' For writing File A data (both original and mapped columns):
            For j = 1 To lastColA
                ' Only process non-mapped columns (we'll add mapped versions right after each original)
                If Right(wsFileA.Cells(3, j).value, 4) <> "_map" Then
                    ' Copy the original column value with format preservation
                    CopyValueWithFormatting wsFileA, i, j, wsResult, rowResult, colOffset
                    colOffset = colOffset + 1
                    
                    ' Check if there's a mapped version and add it
                    mappedColNameA = wsFileA.Cells(3, j).value & "_map"
                    For k = 1 To lastColA
                        If wsFileA.Cells(3, k).value = mappedColNameA Then
                            CopyValueWithFormatting wsFileA, i, k, wsResult, rowResult, colOffset
                            colOffset = colOffset + 1
                            Exit For
                        End If
                    Next k
                End If
            Next j
            
            ' Write supplementary data for File A if needed
            If hasSupplA And suppColsAInitialized Then
                supplMatchRowA = 0
                
                ' Find matching row in supplementary data
                On Error Resume Next
                For j = 4 To UBound(supplDataA, 1)
                    If supplDataA(j, colSupplKeyA) = primaryKeyValueA Then
                        supplMatchRowA = j
                        Exit For
                    End If
                Next j
                On Error GoTo ErrorHandler
                
                ' Write supplementary data
                For j = 0 To UBound(requiredSupplColsA)
                    If requiredSupplColsA(j) <> "" Then
                        If supplMatchRowA > 0 Then
                            On Error Resume Next
                            wsResult.Cells(rowResult, colOffset).value = supplDataA(supplMatchRowA, requiredSupplColIndexA(j))
                            If Err.Number <> 0 Then
                                wsResult.Cells(rowResult, colOffset).value = "Error retrieving data"
                            End If
                            On Error GoTo ErrorHandler
                        Else
                            wsResult.Cells(rowResult, colOffset).value = "N/A - No supplementary record is found"
                        End If
                        colOffset = colOffset + 1
                    End If
                Next j
            End If
            
            ' For writing File B data (both original and mapped columns):
            For j = 1 To lastColB
                ' Only process non-mapped columns
                If Right(wsFileB.Cells(3, j).value, 4) <> "_map" Then
                    ' Copy the original column value with format preservation
                    CopyValueWithFormatting wsFileB, matchRowB, j, wsResult, rowResult, colOffset
                    colOffset = colOffset + 1
                    
                    ' Check if there's a mapped version and add it
                    mappedColNameB = wsFileB.Cells(3, j).value & "_map"
                    For k = 1 To lastColB
                        If wsFileB.Cells(3, k).value = mappedColNameB Then
                            CopyValueWithFormatting wsFileB, matchRowB, k, wsResult, rowResult, colOffset
                            colOffset = colOffset + 1
                            Exit For
                        End If
                    Next k
                End If
            Next j
            
            ' Write supplementary data for File B if needed
            If hasSupplB And suppColsBInitialized Then
                supplMatchRowB = 0
                
                ' Find matching row in supplementary data
                On Error Resume Next
                For j = 4 To UBound(supplDataB, 1)
                    If supplDataB(j, colSupplKeyB) = dataB(matchRowB, colPrimaryKeyB) Then
                        supplMatchRowB = j
                        Exit For
                    End If
                Next j
                On Error GoTo ErrorHandler
                
                ' Write supplementary data
                For j = 0 To UBound(requiredSupplColsB)
                    If requiredSupplColsB(j) <> "" Then
                        If supplMatchRowB > 0 Then
                            On Error Resume Next
                            wsResult.Cells(rowResult, colOffset).value = supplDataB(supplMatchRowB, requiredSupplColIndexB(j))
                            If Err.Number <> 0 Then
                                wsResult.Cells(rowResult, colOffset).value = "Error retrieving data"
                            End If
                            On Error GoTo ErrorHandler
                        Else
                            wsResult.Cells(rowResult, colOffset).value = "N/A - No supplementary record is found"
                        End If
                        colOffset = colOffset + 1
                    End If
                Next j
            End If
            
            ' Write reconciliation results for each column pair
            If mappingCount > 0 Then
                For k = 0 To UBound(mappings, 1)
                    colA = FindColumnIndex(wsFileA, mappings(k, 0))
                    colB = FindColumnIndex(wsFileB, mappings(k, 1))
                    
                    If colA > 0 And colB > 0 Then
                        ' Try to find mapped columns for comparison
                        mappedColNameA = mappings(k, 0) & "_map"
                        mappedColNameB = mappings(k, 1) & "_map"
                        
                        colAComp = FindColumnIndex(wsFileA, mappedColNameA)
                        colBComp = FindColumnIndex(wsFileB, mappedColNameB)
                        
                        ' Use mapped columns if available, otherwise use original
                        If colAComp > 0 Then
                            valueA = wsFileA.Cells(i, colAComp).value
                        Else
                            valueA = wsFileA.Cells(i, colA).value
                        End If
                        
                        If colBComp > 0 Then
                            valueB = wsFileB.Cells(matchRowB, colBComp).value
                        Else
                            valueB = wsFileB.Cells(matchRowB, colB).value
                        End If
                        
                        If valueA = valueB Then
                            wsResult.Cells(rowResult, colOffset).value = "MATCH"
                            wsResult.Cells(rowResult, colOffset).Interior.Color = RGB(198, 239, 206) ' Light green
                        Else
                            wsResult.Cells(rowResult, colOffset).value = "UNMATCH"
                            wsResult.Cells(rowResult, colOffset).Interior.Color = RGB(255, 199, 206) ' Light red
                        End If
                        
                        colOffset = colOffset + 1
                    End If
                Next k
            End If
        Else
            ' This is the ELSE branch for if no match is found (singleton in A)
            wsResult.Cells(rowResult, 1).value = "Singleton in " & gstrInputASource
            
            ' Write File A data to Result sheet - both original and mapped columns
            colOffset = 2 ' Start after Recon Result column
            
            ' For writing File A data (both original and mapped columns):
            For j = 1 To lastColA
                ' Only process non-mapped columns (we'll add mapped versions right after each original)
                If Right(wsFileA.Cells(3, j).value, 4) <> "_map" Then
                    ' Copy the original column value with format preservation
                    CopyValueWithFormatting wsFileA, i, j, wsResult, rowResult, colOffset
                    colOffset = colOffset + 1
                    
                    ' Check if there's a mapped version and add it
                    mappedColNameA = wsFileA.Cells(3, j).value & "_map"
                    For k = 1 To lastColA
                        If wsFileA.Cells(3, k).value = mappedColNameA Then
                            CopyValueWithFormatting wsFileA, i, k, wsResult, rowResult, colOffset
                            colOffset = colOffset + 1
                            Exit For
                        End If
                    Next k
                End If
            Next j
            
            ' Write supplementary data for File A if needed (for singletons)
            If hasSupplA And suppColsAInitialized Then
                supplMatchRowA = 0
                
                ' Find matching row in supplementary data
                On Error Resume Next
                For j = 4 To UBound(supplDataA, 1)
                    If supplDataA(j, colSupplKeyA) = primaryKeyValueA Then
                        supplMatchRowA = j
                        Exit For
                    End If
                Next j
                On Error GoTo ErrorHandler
                
                ' Write supplementary data
                For j = 0 To UBound(requiredSupplColsA)
                    If requiredSupplColsA(j) <> "" Then
                        If supplMatchRowA > 0 Then
                            On Error Resume Next
                            wsResult.Cells(rowResult, colOffset).value = supplDataA(supplMatchRowA, requiredSupplColIndexA(j))
                            If Err.Number <> 0 Then
                                wsResult.Cells(rowResult, colOffset).value = "Error retrieving data"
                            End If
                            On Error GoTo ErrorHandler
                        Else
                            wsResult.Cells(rowResult, colOffset).value = "N/A - No supplementary record is found"
                        End If
                        colOffset = colOffset + 1
                    End If
                Next j
            End If
            
            ' For singleton in A, leave B columns empty
            For j = 1 To lastColB
                ' Count both original and mapped columns to skip
                If Right(wsFileB.Cells(3, j).value, 4) <> "_map" Then
                    colOffset = colOffset + 1 ' Skip original column
                    
                    ' Check if there's a mapped version and skip it too
                    mappedColNameB = wsFileB.Cells(3, j).value & "_map"
                    For k = 1 To lastColB
                        If wsFileB.Cells(3, k).value = mappedColNameB Then
                            colOffset = colOffset + 1 ' Skip mapped column
                            Exit For
                        End If
                    Next k
                End If
            Next j
            
            ' Skip supplementary B columns if needed
            If hasSupplB And suppColsBInitialized Then
                Dim suppBColCount As Long
                suppBColCount = 0
                
                ' Count non-empty supplementary columns
                For j = 0 To UBound(requiredSupplColsB)
                    If requiredSupplColsB(j) <> "" Then
                        suppBColCount = suppBColCount + 1
                    End If
                Next j
                
                colOffset = colOffset + suppBColCount
            End If
            
            ' Mark recon columns as Singleton
            For j = 1 To reconColCount
                wsResult.Cells(rowResult, colOffset).value = "Singleton"
                wsResult.Cells(rowResult, colOffset).Interior.Color = RGB(255, 235, 156) ' Light orange
                colOffset = colOffset + 1
            Next j
        End If
        
        ' Move to next result row
        rowResult = rowResult + 1
    Next i
    
    ' Find singletons in File B (records in B with no match in A)
    For i = 4 To lastRowB
        ' Get key values
        primaryKeyValueB = dataB(i, colPrimaryKeyB)
        
        If colSecondaryKeyB > 0 Then
            secondaryKeyValueB = dataB(i, colSecondaryKeyB)
        Else
            secondaryKeyValueB = ""
        End If
        
        ' Check if this record has a match in File A
        matchFound = False
        
        ' Try to match with both primary and secondary keys if available
        If colSecondaryKeyA > 0 And colSecondaryKeyB > 0 Then
            For j = 4 To lastRowA
                If dataA(j, colPrimaryKeyA) = primaryKeyValueB And _
                   dataA(j, colSecondaryKeyA) = secondaryKeyValueB Then
                    matchFound = True
                    Exit For
                End If
            Next j
        End If
        
        ' If no match found with both keys, try primary key only
        If Not matchFound Then
            For j = 4 To lastRowA
                If dataA(j, colPrimaryKeyA) = primaryKeyValueB Then
                    matchFound = True
                    Exit For
                End If
            Next j
        End If
        
        ' If no match found, this is a singleton in B
        If Not matchFound Then
            ' Recon Result column
            wsResult.Cells(rowResult, 1).value = "Singleton in " & gstrInputBSource
            
            ' Skip File A columns (both original and mapped)
            colOffset = 2 ' Start after Recon Result column
            
            For j = 1 To lastColA
                ' Only process non-mapped columns
                If Right(wsFileA.Cells(3, j).value, 4) <> "_map" Then
                    colOffset = colOffset + 1 ' Skip original column
                    
                    ' Check if there's a mapped version and skip it too
                    mappedColNameA = wsFileA.Cells(3, j).value & "_map"
                    For k = 1 To lastColA
                        If wsFileA.Cells(3, k).value = mappedColNameA Then
                            colOffset = colOffset + 1 ' Skip mapped column
                            Exit For
                        End If
                    Next k
                End If
            Next j
            
            ' Skip supplementary A columns if needed
            If hasSupplA And suppColsAInitialized Then
                Dim suppAColCount As Long
                suppAColCount = 0
                
                ' Count non-empty supplementary columns
                For j = 0 To UBound(requiredSupplColsA)
                    If requiredSupplColsA(j) <> "" Then
                        suppAColCount = suppAColCount + 1
                    End If
                Next j
                
                colOffset = colOffset + suppAColCount
            End If
            
            ' Write File B data - both original and mapped
            For j = 1 To lastColB
                ' Only process non-mapped columns
                If Right(wsFileB.Cells(3, j).value, 4) <> "_map" Then
                    ' Copy the original column value
                    CopyValueWithFormatting wsFileB, i, j, wsResult, rowResult, colOffset
                    colOffset = colOffset + 1
                    
                    ' Check if there's a mapped version and add it
                    mappedColNameB = wsFileB.Cells(3, j).value & "_map"
                    For k = 1 To lastColB
                        If wsFileB.Cells(3, k).value = mappedColNameB Then
                            CopyValueWithFormatting wsFileB, i, k, wsResult, rowResult, colOffset
                            colOffset = colOffset + 1
                            Exit For
                        End If
                    Next k
                End If
            Next j
            
            ' Write supplementary data for File B if needed
            If hasSupplB And suppColsBInitialized Then
                singletonSupplMatchRowB = 0
                
                ' Find matching row in supplementary data
                On Error Resume Next
                For j = 4 To UBound(supplDataB, 1)
                    If supplDataB(j, colSupplKeyB) = primaryKeyValueB Then
                        singletonSupplMatchRowB = j
                        Exit For
                    End If
                Next j
                On Error GoTo ErrorHandler
                
                ' Write supplementary data
                For j = 0 To UBound(requiredSupplColsB)
                    If requiredSupplColsB(j) <> "" Then
                        If singletonSupplMatchRowB > 0 Then
                            On Error Resume Next
                            wsResult.Cells(rowResult, colOffset).value = supplDataB(singletonSupplMatchRowB, requiredSupplColIndexB(j))
                            If Err.Number <> 0 Then
                                wsResult.Cells(rowResult, colOffset).value = "Error retrieving data"
                            End If
                            On Error GoTo ErrorHandler
                        Else
                            wsResult.Cells(rowResult, colOffset).value = "N/A - No supplementary record is found"
                        End If
                        colOffset = colOffset + 1
                    End If
                Next j
            End If
            
            ' Mark recon columns as Singleton
            For j = 1 To reconColCount
                wsResult.Cells(rowResult, colOffset).value = "Singleton"
                wsResult.Cells(rowResult, colOffset).Interior.Color = RGB(255, 235, 156) ' Light orange
                colOffset = colOffset + 1
            Next j
            
            ' Move to next result row
            rowResult = rowResult + 1
        End If
    Next i
    
    ' Format the Result sheet
    With wsResult.Range("A3:ZZ3")
        .Font.Bold = True
        .Borders.LineStyle = xlContinuous
    End With
    
    With wsResult.Range("A3:ZZ" & (rowResult - 1))
        .Borders.LineStyle = xlContinuous
    End With
    
    ' Auto-fit columns
    wsResult.Columns("B:ZZ").AutoFit
    
    Exit Sub
    
ErrorHandler:
    ' Show detailed error information
    MsgBox "Error " & Err.Number & " (" & Err.Description & ")" & vbCrLf & _
           "Procedure: PerformReconciliation" & vbCrLf & _
           "Line: " & Erl, vbCritical, "Error"
End Sub
Function FindColumnIndex(ws As Worksheet, columnName As String) As Long
    ' Find the index of a column with the specified name in row 3
    Dim lastCol As Long
    Dim i As Long
    
    FindColumnIndex = 0 ' Default to not found
    
    If columnName = "" Then Exit Function
    
    lastCol = ws.Cells(3, ws.Columns.count).End(xlToLeft).Column
    
    For i = 1 To lastCol
        If ws.Cells(3, i).value = columnName Then
            FindColumnIndex = i
            Exit Function
        End If
    Next i
End Function

Function GetRequiredSupplementaryColumns(fileType As String) As String()
    ' Get the list of required supplementary columns from the Input Parameters sheet
    Dim ws As Worksheet
    Dim startRow As Long
    Dim i As Long
    Dim colCount As Long
    Dim tempArray() As String
    
    ' Set worksheet
    Set ws = Worksheets("Input Parameters")
    
    ' Determine the start row based on file type
    If fileType = "A" Then
        startRow = 35 ' Adjust if needed based on your actual structure
    ElseIf fileType = "B" Then
        startRow = 46 ' Adjust if needed based on your actual structure
    Else
        ReDim tempArray(0)
        GetRequiredSupplementaryColumns = tempArray
        Exit Function
    End If
    
    ' Count how many columns are required
    colCount = 0
    i = startRow
    
    Do While ws.Cells(i, 1).value <> ""
        colCount = colCount + 1
        i = i + 1
    Loop
    
    ' If no columns required, return empty array
    If colCount = 0 Then
        ReDim tempArray(0)
        GetRequiredSupplementaryColumns = tempArray
        Exit Function
    End If
    
    ' Get the column names
    ReDim tempArray(colCount - 1)
    
    For i = 0 To colCount - 1
        tempArray(i) = ws.Cells(startRow + i, 1).value
    Next i
    
    GetRequiredSupplementaryColumns = tempArray
End Function

' =================================== DASHBOARD FUNCTIONS ===================================

Sub ClearDashboardSheet()
    ' Clear both content and formatting from the Dashboard sheet
    On Error GoTo ErrorHandler
    
    Dim wsDashboard As Worksheet
    
    ' Set worksheet
    Set wsDashboard = Worksheets("Dashboard")
    
    ' Preserve title and subtitle
    Dim title As String, subtitle As String
    title = wsDashboard.Range("A1").value
    subtitle = wsDashboard.Range("A2").value
    
    ' Clear everything from row 3 down (both content and formatting)
    wsDashboard.Range("A3:ZZ" & wsDashboard.Rows.count).ClearContents
    wsDashboard.Range("A3:ZZ" & wsDashboard.Rows.count).ClearFormats
    
    ' Delete any existing charts
    On Error Resume Next
    If wsDashboard.ChartObjects.count > 0 Then
        wsDashboard.ChartObjects.Delete
    End If
    On Error GoTo ErrorHandler
    
    ' Restore title and subtitle
    wsDashboard.Range("A1").value = title
    wsDashboard.Range("A1").Font.Bold = True
    wsDashboard.Range("A1").Font.Size = 16
    wsDashboard.Range("A2").value = subtitle
    wsDashboard.Range("A2").Font.Italic = True
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in ClearDashboardSheet", vbCritical, "Error"
End Sub

Sub GenerateDashboard()
    ' Generate dashboard with reconciliation summary
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    
    Dim wsDashboard As Worksheet
    Dim wsResult As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim totalCount As Long
    Dim matchCount As Long
    Dim unmatchCount As Long
    Dim singletonACount As Long
    Dim singletonBCount As Long
    Dim reconResult As String
    
    ' Set worksheets
    Set wsDashboard = Worksheets("Dashboard")
    Set wsResult = Worksheets("Result")
    
    ' Clear dashboard (both content and formatting)
    ClearDashboardSheet
    
    ' Check if reconciliation has been run
    lastRow = wsResult.Cells(wsResult.Rows.count, 1).End(xlUp).Row
    
    If lastRow <= 3 Then
        MsgBox "No reconciliation results found. Please run reconciliation first.", vbExclamation, "No Data"
        GoTo CleanExit
    End If
    
    ' Count reconciliation statuses
    totalCount = lastRow - 3
    matchCount = 0
    unmatchCount = 0
    singletonACount = 0
    singletonBCount = 0
    
    For i = 4 To lastRow
        reconResult = wsResult.Cells(i, 1).value
        
        If InStr(1, reconResult, "Full Match", vbTextCompare) > 0 Or _
           InStr(1, reconResult, "Primary Key Match", vbTextCompare) > 0 Then
            ' Check if contains "All Fields Match"
            If InStr(1, reconResult, "All Fields Match", vbTextCompare) > 0 Then
                matchCount = matchCount + 1
            Else
                unmatchCount = unmatchCount + 1
            End If
        ElseIf InStr(1, reconResult, "Singleton in " & gstrInputASource, vbTextCompare) > 0 Then
            singletonACount = singletonACount + 1
        ElseIf InStr(1, reconResult, "Singleton in " & gstrInputBSource, vbTextCompare) > 0 Then
            singletonBCount = singletonBCount + 1
        End If
    Next i
    
    ' Populate dashboard summary
    wsDashboard.Range("A3").value = "RECONCILIATION SUMMARY"
    wsDashboard.Range("A3").Font.Bold = True
    wsDashboard.Range("A3").Font.Size = 14
    
    wsDashboard.Range("A5").value = "Total Records Processed:"
    wsDashboard.Range("B5").value = totalCount
    
    wsDashboard.Range("A6").value = "Fully Matched Records:"
    wsDashboard.Range("B6").value = matchCount
    wsDashboard.Range("C6").value = Format(matchCount / totalCount, "0.0%")
    
    wsDashboard.Range("A7").value = "Records with Differences:"
    wsDashboard.Range("B7").value = unmatchCount
    wsDashboard.Range("C7").value = Format(unmatchCount / totalCount, "0.0%")
    
    wsDashboard.Range("A8").value = "Singletons in " & gstrInputASource & ":"
    wsDashboard.Range("B8").value = singletonACount
    wsDashboard.Range("C8").value = Format(singletonACount / totalCount, "0.0%")
    
    wsDashboard.Range("A9").value = "Singletons in " & gstrInputBSource & ":"
    wsDashboard.Range("B9").value = singletonBCount
    wsDashboard.Range("C9").value = Format(singletonBCount / totalCount, "0.0%")
    
    ' Format dashboard
    wsDashboard.Range("A5:C9").Borders.LineStyle = xlContinuous
    
    ' Add a simple chart
    CreateDashboardChart wsDashboard, matchCount, unmatchCount, singletonACount, singletonBCount
    
    ' Auto-fit columns
    wsDashboard.Columns("A:C").AutoFit
    
CleanExit:
    Application.ScreenUpdating = True
    
    If Err.Number = 0 Then
        MsgBox "Dashboard generated successfully!", vbInformation, "Dashboard Complete"
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in GenerateDashboard", vbCritical, "Error"
    Resume CleanExit
End Sub

Function IsRowFullMatch(wsResult As Worksheet, rowNum As Long) As Boolean
    ' Check if all reconciliation columns in a row show "MATCH"
    Dim lastCol As Long
    Dim i As Long
    
    lastCol = wsResult.Cells(3, wsResult.Columns.count).End(xlToLeft).Column
    
    ' Start from the first reconciliation column
    For i = lastCol - 10 To lastCol ' Assumption: last 10 columns are recon results
        If wsResult.Cells(3, i).value Like "*Recon*" Then
            If wsResult.Cells(rowNum, i).value <> "MATCH" And _
               wsResult.Cells(rowNum, i).value <> "Singleton" Then
                IsRowFullMatch = False
                Exit Function
            End If
        End If
    Next i
    
    IsRowFullMatch = True
End Function

Sub CreateDashboardChart(ws As Worksheet, matchCount As Long, unmatchCount As Long, _
                         singletonACount As Long, singletonBCount As Long)
    ' Create a simple chart in the dashboard
    Dim chartObj As ChartObject
    Dim chartData As Range
    
    ' Create data for chart
    ws.Range("E5").value = "Category"
    ws.Range("F5").value = "Count"
    
    ws.Range("E6").value = "Matched"
    ws.Range("F6").value = matchCount
    
    ws.Range("E7").value = "Unmatched"
    ws.Range("F7").value = unmatchCount
    
    ws.Range("E8").value = "Singleton in " & gstrInputASource
    ws.Range("F8").value = singletonACount
    
    ws.Range("E9").value = "Singleton in " & gstrInputBSource
    ws.Range("F9").value = singletonBCount
    
    ' Create chart
    Set chartData = ws.Range("E5:F9")
    
    ' Delete any existing charts
    If ws.ChartObjects.count > 0 Then
        ws.ChartObjects.Delete
    End If
    
    ' Create new chart
    Set chartObj = ws.ChartObjects.Add(320, 50, 400, 250)
    
    With chartObj.Chart
        .SetSourceData Source:=chartData
        .ChartType = xlPie
        .HasTitle = True
        .ChartTitle.Text = "Reconciliation Results"
        .HasLegend = True
        .Legend.Position = xlLegendPositionRight
    End With
    
    ' Apply colors
    chartObj.Chart.SeriesCollection(1).Points(1).Interior.Color = RGB(146, 208, 80) ' Green for matches
    chartObj.Chart.SeriesCollection(1).Points(2).Interior.Color = RGB(255, 0, 0)    ' Red for unmatches
    chartObj.Chart.SeriesCollection(1).Points(3).Interior.Color = RGB(255, 192, 0)  ' Orange for singleton A
    chartObj.Chart.SeriesCollection(1).Points(4).Interior.Color = RGB(0, 176, 240)  ' Blue for singleton B
End Sub

' =================================== RUN ALL STEPS ===================================

Sub RunAllSteps()
    ' Run all reconciliation steps in sequence
    Application.ScreenUpdating = False
    
    ' Load Input Files
    LoadInputFiles
    
    ' Load Supplementary Files
    LoadSupplementaryFiles
    
    ' Apply Mapping
    ApplyMapping
    
    ' Run Reconciliation
    RunReconciliation
    
    ' Generate Dashboard
    GenerateDashboard
    
    Application.ScreenUpdating = True
    MsgBox "All reconciliation steps completed successfully!", vbInformation, "Process Complete"
End Sub

' =================================== PLATFORM SPECIFIC NOTES ===================================

' ===== WINDOWS VS MACOS COMPATIBILITY NOTES =====
'
' 1. File Paths:
'    - macOS uses forward slashes (/) in file paths
'    - Windows uses backslashes (\) in file paths
'    - When moving to Windows, you may need to update file paths or use Replace() to convert slashes
'
' 2. File Dialogs:
'    - This code uses simple file path input via cells rather than file dialogs for better cross-platform compatibility
'    - If you add file dialogs later, note that Application.FileDialog works differently across platforms
'
' 3. ActiveX Controls:
'    - This code uses Form controls (Buttons) instead of ActiveX controls for better macOS compatibility
'    - On Windows, you might prefer ActiveX controls for more formatting options
'
' 4. Line Endings:
'    - macOS and Windows use different line ending characters, which can affect text file parsing
'    - The code handles this automatically using Line Input
'
' 5. Performance:
'    - Excel on Windows typically performs better with large datasets
'    - If you have performance issues on macOS, consider loading smaller chunks of data
'
' 6. Color Models:
'    - This code uses RGB() for colors which works on both platforms
'    - Some advanced color features might behave differently across platforms
'
' 7. Chart Objects:
'    - Chart object behavior can vary slightly between platforms
'    - The code uses basic chart properties that work on both platforms

