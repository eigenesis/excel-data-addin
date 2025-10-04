Attribute VB_Name = "DataExtractInsert"
' Data Extract & Insert Tool
' Excel VBA Macros for extracting and inserting data

' Global variable to store extracted data
Public ExtractedData As Variant

' ====================================
' EXTRACT DATA FUNCTIONS
' ====================================

Sub ExtractFromCurrentSheet()
    Dim ws As Worksheet
    Dim usedRange As Range

    Set ws = ActiveSheet
    Set usedRange = ws.UsedRange

    ExtractedData = usedRange.Value

    MsgBox "Extracted " & usedRange.Rows.Count & " rows x " & usedRange.Columns.Count & " columns" & vbCrLf & _
           "From: " & ws.Name & vbCrLf & vbCrLf & _
           "Data is now ready to insert elsewhere.", vbInformation, "Data Extracted"
End Sub

Sub ExtractFromSpecificSheet()
    Dim sheetName As String
    Dim ws As Worksheet
    Dim usedRange As Range

    sheetName = InputBox("Enter the sheet name to extract from:", "Extract from Specific Sheet")

    If sheetName = "" Then Exit Sub

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "Sheet '" & sheetName & "' not found!", vbExclamation
        Exit Sub
    End If

    Set usedRange = ws.UsedRange
    ExtractedData = usedRange.Value

    MsgBox "Extracted " & usedRange.Rows.Count & " rows x " & usedRange.Columns.Count & " columns" & vbCrLf & _
           "From: " & ws.Name & vbCrLf & vbCrLf & _
           "Data is now ready to insert elsewhere.", vbInformation, "Data Extracted"
End Sub

Sub ExtractFromRange()
    Dim rangeAddr As String
    Dim targetRange As Range

    rangeAddr = InputBox("Enter the range to extract (e.g., A1:C10):", "Extract from Range")

    If rangeAddr = "" Then Exit Sub

    On Error Resume Next
    Set targetRange = ActiveSheet.Range(rangeAddr)
    On Error GoTo 0

    If targetRange Is Nothing Then
        MsgBox "Invalid range address!", vbExclamation
        Exit Sub
    End If

    ExtractedData = targetRange.Value

    MsgBox "Extracted " & targetRange.Rows.Count & " rows x " & targetRange.Columns.Count & " columns" & vbCrLf & _
           "From: " & targetRange.Address & vbCrLf & vbCrLf & _
           "Data is now ready to insert elsewhere.", vbInformation, "Data Extracted"
End Sub

Sub ExtractFromCSVFile()
    Dim filePath As Variant
    Dim fileNum As Integer
    Dim lineText As String
    Dim dataArray() As String
    Dim rowCount As Long
    Dim i As Long

    ' Open file dialog
    filePath = Application.GetOpenFilename("CSV Files (*.csv), *.csv", , "Select CSV File to Import")

    If filePath = False Then Exit Sub

    ' Read CSV file
    fileNum = FreeFile
    Open filePath For Input As fileNum

    rowCount = 0
    ReDim dataArray(0)

    Do Until EOF(fileNum)
        Line Input #fileNum, lineText
        rowCount = rowCount + 1
        ReDim Preserve dataArray(rowCount - 1)
        dataArray(rowCount - 1) = lineText
    Loop

    Close fileNum

    ' Parse CSV into array
    ExtractedData = ParseCSV(dataArray, rowCount)

    MsgBox "Extracted " & rowCount & " rows from CSV file" & vbCrLf & vbCrLf & _
           "Data is now ready to insert.", vbInformation, "CSV Imported"
End Sub

' Helper function to parse CSV
Private Function ParseCSV(lines() As String, rowCount As Long) As Variant
    Dim result() As Variant
    Dim i As Long, j As Long
    Dim fields() As String
    Dim maxCols As Long

    ' First pass: determine max columns
    maxCols = 0
    For i = 0 To rowCount - 1
        fields = Split(lines(i), ",")
        If UBound(fields) + 1 > maxCols Then maxCols = UBound(fields) + 1
    Next i

    ' Second pass: populate array
    ReDim result(1 To rowCount, 1 To maxCols)

    For i = 0 To rowCount - 1
        fields = Split(lines(i), ",")
        For j = 0 To UBound(fields)
            result(i + 1, j + 1) = Trim(fields(j))
        Next j
    Next i

    ParseCSV = result
End Function

' ====================================
' INSERT DATA FUNCTIONS
' ====================================

Sub InsertAtCurrentSelection()
    If IsEmpty(ExtractedData) Then
        MsgBox "No data to insert! Please extract data first.", vbExclamation, "No Data"
        Exit Sub
    End If

    Dim targetRange As Range
    Dim rowsCount As Long
    Dim colsCount As Long

    ' Determine dimensions
    If IsArray(ExtractedData) Then
        rowsCount = UBound(ExtractedData, 1) - LBound(ExtractedData, 1) + 1
        colsCount = UBound(ExtractedData, 2) - LBound(ExtractedData, 2) + 1
    Else
        rowsCount = 1
        colsCount = 1
    End If

    Set targetRange = Selection.Resize(rowsCount, colsCount)
    targetRange.Value = ExtractedData

    MsgBox "Inserted " & rowsCount & " rows x " & colsCount & " columns" & vbCrLf & _
           "At: " & targetRange.Address, vbInformation, "Data Inserted"
End Sub

Sub InsertInNewSheet()
    If IsEmpty(ExtractedData) Then
        MsgBox "No data to insert! Please extract data first.", vbExclamation, "No Data"
        Exit Sub
    End If

    Dim newSheet As Worksheet
    Dim targetRange As Range
    Dim rowsCount As Long
    Dim colsCount As Long

    ' Create new sheet
    Set newSheet = ThisWorkbook.Worksheets.Add
    newSheet.Name = "ImportedData_" & Format(Now, "hhmmss")

    ' Determine dimensions
    If IsArray(ExtractedData) Then
        rowsCount = UBound(ExtractedData, 1) - LBound(ExtractedData, 1) + 1
        colsCount = UBound(ExtractedData, 2) - LBound(ExtractedData, 2) + 1
    Else
        rowsCount = 1
        colsCount = 1
    End If

    Set targetRange = newSheet.Range("A1").Resize(rowsCount, colsCount)
    targetRange.Value = ExtractedData

    ' Auto-fit columns
    newSheet.Columns.AutoFit

    MsgBox "Inserted " & rowsCount & " rows x " & colsCount & " columns" & vbCrLf & _
           "In new sheet: " & newSheet.Name, vbInformation, "Data Inserted"
End Sub

Sub InsertAtSpecificRange()
    If IsEmpty(ExtractedData) Then
        MsgBox "No data to insert! Please extract data first.", vbExclamation, "No Data"
        Exit Sub
    End If

    Dim rangeAddr As String
    Dim targetRange As Range
    Dim rowsCount As Long
    Dim colsCount As Long

    rangeAddr = InputBox("Enter the starting cell for insertion (e.g., A1):", "Insert at Specific Range")

    If rangeAddr = "" Then Exit Sub

    On Error Resume Next
    Set targetRange = ActiveSheet.Range(rangeAddr)
    On Error GoTo 0

    If targetRange Is Nothing Then
        MsgBox "Invalid range address!", vbExclamation
        Exit Sub
    End If

    ' Determine dimensions
    If IsArray(ExtractedData) Then
        rowsCount = UBound(ExtractedData, 1) - LBound(ExtractedData, 1) + 1
        colsCount = UBound(ExtractedData, 2) - LBound(ExtractedData, 2) + 1
    Else
        rowsCount = 1
        colsCount = 1
    End If

    Set targetRange = targetRange.Resize(rowsCount, colsCount)
    targetRange.Value = ExtractedData

    MsgBox "Inserted " & rowsCount & " rows x " & colsCount & " columns" & vbCrLf & _
           "At: " & targetRange.Address, vbInformation, "Data Inserted"
End Sub

Sub InsertAsTable()
    If IsEmpty(ExtractedData) Then
        MsgBox "No data to insert! Please extract data first.", vbExclamation, "No Data"
        Exit Sub
    End If

    Dim targetRange As Range
    Dim rowsCount As Long
    Dim colsCount As Long
    Dim tbl As ListObject

    ' Determine dimensions
    If IsArray(ExtractedData) Then
        rowsCount = UBound(ExtractedData, 1) - LBound(ExtractedData, 1) + 1
        colsCount = UBound(ExtractedData, 2) - LBound(ExtractedData, 2) + 1
    Else
        rowsCount = 1
        colsCount = 1
    End If

    Set targetRange = Selection.Resize(rowsCount, colsCount)
    targetRange.Value = ExtractedData

    ' Create table
    Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, targetRange, , xlYes)
    tbl.Name = "DataTable_" & Format(Now, "hhmmss")
    tbl.TableStyle = "TableStyleMedium2"

    MsgBox "Inserted " & rowsCount & " rows x " & colsCount & " columns as table" & vbCrLf & _
           "Table name: " & tbl.Name, vbInformation, "Table Created"
End Sub

' ====================================
' SHOW MENU
' ====================================

Sub ShowDataToolsMenu()
    Dim response As VbMsgBoxResult

    response = MsgBox("Data Extract & Insert Tool" & vbCrLf & vbCrLf & _
                     "Choose an action:" & vbCrLf & vbCrLf & _
                     "YES = Extract Data" & vbCrLf & _
                     "NO = Insert Data" & vbCrLf & _
                     "CANCEL = Exit", vbYesNoCancel + vbQuestion, "Data Tools")

    If response = vbYes Then
        ShowExtractMenu
    ElseIf response = vbNo Then
        ShowInsertMenu
    End If
End Sub

Sub ShowExtractMenu()
    Dim choice As Integer

    choice = Application.InputBox("EXTRACT DATA - Choose source:" & vbCrLf & vbCrLf & _
                                 "1 = Current Sheet" & vbCrLf & _
                                 "2 = Specific Sheet" & vbCrLf & _
                                 "3 = Range" & vbCrLf & _
                                 "4 = CSV File", "Extract Data", 1, Type:=1)

    Select Case choice
        Case 1: ExtractFromCurrentSheet
        Case 2: ExtractFromSpecificSheet
        Case 3: ExtractFromRange
        Case 4: ExtractFromCSVFile
    End Select
End Sub

Sub ShowInsertMenu()
    Dim choice As Integer

    choice = Application.InputBox("INSERT DATA - Choose destination:" & vbCrLf & vbCrLf & _
                                 "1 = Current Selection" & vbCrLf & _
                                 "2 = New Sheet" & vbCrLf & _
                                 "3 = Specific Range" & vbCrLf & _
                                 "4 = As Formatted Table", "Insert Data", 1, Type:=1)

    Select Case choice
        Case 1: InsertAtCurrentSelection
        Case 2: InsertInNewSheet
        Case 3: InsertAtSpecificRange
        Case 4: InsertAsTable
    End Select
End Sub