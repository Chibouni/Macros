Sub ExtractAndCenterTestData_v5()
    Dim wordDoc As Document
    Dim excelApp As Object
    Dim excelWb As Object
    Dim excelWs As Object
    Dim regExID As Object, regExIssue As Object
    Dim matches As Object
    
    Dim tbl As Table
    Dim rHeader As Long, r As Long, c As Long
    Dim colStatus As Long, colComments As Long, colIssue As Long
    Dim txt As String, cleanText As String
    
    Dim currentTestID As String
    Dim currentStatus As String
    Dim currentIssue As String
    Dim currentComment As String
    
    ' Excel Table trackers
    Dim nokRow As Long, okRow As Long, npRow As Long, emptyRow As Long
    
    Set wordDoc = ActiveDocument
    
    ' Setup Regex patterns
    Set regExID = CreateObject("VBScript.RegExp")
    regExID.Pattern = "SWTP_[A-Z0-9_]+"
    regExID.IgnoreCase = True
    
    Set regExIssue = CreateObject("VBScript.RegExp")
    regExIssue.Pattern = "CYBSAF-\d+"
    regExIssue.IgnoreCase = True
    
    ' Initialize Excel Destination
    On Error Resume Next
    Set excelApp = GetObject(, "Excel.Application")
    If excelApp Is Nothing Then
        Set excelApp = CreateObject("Excel.Application")
    End If
    On Error GoTo 0
    
    excelApp.Visible = True
    Set excelWb = excelApp.Workbooks.Add
    Set excelWs = excelWb.Sheets(1)
    excelWs.Name = "Test Dashboard"
    excelApp.ActiveWindow.DisplayGridLines = True
    
    ' ==========================================
    ' 1. EXCEL STATISTICS / HEADERS SETUP
    ' ==========================================
    excelWs.Cells(2, 1).Value = "TEST EXECUTION SUMMARY"
    excelWs.Cells(2, 1).Font.Size = 12
    excelWs.Cells(2, 1).Font.Bold = True
    excelWs.Cells(2, 1).Font.Color = RGB(26, 54, 93)
    
    excelWs.Cells(4, 1).Value = "Metric"
    excelWs.Cells(4, 2).Value = "Count"
    excelWs.Range("A4:B4").Font.Bold = True
    excelWs.Range("A4:B4").Interior.Color = RGB(226, 232, 240)
    
    excelWs.Cells(5, 1).Value = "Passed Tests (OK)"
    excelWs.Cells(6, 1).Value = "Failed Tests (NOK)"
    excelWs.Cells(7, 1).Value = "Not Processed (NP)"
    excelWs.Cells(8, 1).Value = "Missing Status (EMPTY)"
    excelWs.Cells(9, 1).Value = "Total Tests"
    excelWs.Cells(9, 1).Font.Bold = True
    
    excelWs.Cells(5, 2).Value = "=COUNTA(E13:E5000)"
    excelWs.Cells(6, 2).Value = "=COUNTA(A13:A5000)"
    excelWs.Cells(7, 2).Value = "=COUNTA(H13:H5000)"
    excelWs.Cells(8, 2).Value = "=COUNTA(M13:M5000)"
    excelWs.Cells(9, 2).Value = "=SUM(B5:B8)"
    excelWs.Cells(9, 2).Font.Bold = True
    excelWs.Range("A4:B9").Borders.LineStyle = 1
    
    Dim startRow As Long
    startRow = 11
    
    ' Setup Table Headers in Excel
    excelWs.Cells(startRow, 1).Value = "Failed Tests (NOK)"
    excelWs.Range("A12:C12").Font.Bold = True
    excelWs.Cells(startRow + 1, 1).Value = "Test ID"
    excelWs.Cells(startRow + 1, 2).Value = "Status"
    excelWs.Cells(startRow + 1, 3).Value = "Associated Issue"
    excelWs.Range("A12:C12").Interior.Color = RGB(140, 45, 25)
    excelWs.Range("A12:C12").Font.Color = RGB(255, 255, 255)
    excelWs.Range("A12:C12").HorizontalAlignment = -4108
    nokRow = startRow + 2
    
    excelWs.Cells(startRow, 5).Value = "Passed Tests (OK)"
    excelWs.Range("E12:F12").Font.Bold = True
    excelWs.Cells(startRow + 1, 5).Value = "Test ID"
    excelWs.Cells(startRow + 1, 6).Value = "Status"
    excelWs.Range("E12:F12").Interior.Color = RGB(46, 105, 48)
    excelWs.Range("E12:F12").Font.Color = RGB(255, 255, 255)
    excelWs.Range("E12:F12").HorizontalAlignment = -4108
    okRow = startRow + 2
    
    excelWs.Cells(startRow, 8).Value = "Not Processed (NP)"
    excelWs.Range("H12:K12").Font.Bold = True
    excelWs.Cells(startRow + 1, 8).Value = "Test ID"
    excelWs.Cells(startRow + 1, 9).Value = "Status"
    excelWs.Cells(startRow + 1, 10).Value = "Associated Issue"
    excelWs.Cells(startRow + 1, 11).Value = "Comments (No Issue)"
    excelWs.Range("H12:K12").Interior.Color = RGB(74, 101, 114)
    excelWs.Range("H12:K12").Font.Color = RGB(255, 255, 255)
    excelWs.Range("H12:K12").HorizontalAlignment = -4108
    npRow = startRow + 2
    
    excelWs.Cells(startRow, 13).Value = "Missing Status (EMPTY)"
    excelWs.Range("M12:N12").Font.Bold = True
    excelWs.Cells(startRow + 1, 13).Value = "Test ID"
    excelWs.Cells(startRow + 1, 14).Value = "Status"
    excelWs.Range("M12:N12").Interior.Color = RGB(218, 165, 32)
    excelWs.Range("M12:N12").Font.Color = RGB(255, 255, 255)
    excelWs.Range("M12:N12").HorizontalAlignment = -4108
    emptyRow = startRow + 2

    ' ==========================================
    ' 2. LOOP WORD TABLES, APPLY LOGIC, CENTERING
    ' ==========================================
    For Each tbl In wordDoc.Tables
        colStatus = 0
        colComments = 0
        colIssue = 0
        rHeader = 0
        currentTestID = ""
        
        ' Find the parent tracking Test ID by checking text right above this table structure
        On Error Resume Next
        cleanText = tbl.Range.Previous(Unit:=wdParagraph, Count:=1).Text
        If regExID.Test(cleanText) Then
            Set matches = regExID.Execute(cleanText)
            currentTestID = matches(0).Value
        Else
            Dim pCount As Long
            For pCount = 2 To 4
                cleanText = tbl.Range.Previous(Unit:=wdParagraph, Count:=pCount).Text
                If regExID.Test(cleanText) Then
                    Set matches = regExID.Execute(cleanText)
                    currentTestID = matches(0).Value
                    Exit For
                End If
            Next pCount
        End If
        On Error GoTo 0
        
        ' Map columns based on your Rectify logic
        For r = 1 To tbl.Rows.Count
            For c = 1 To tbl.Columns.Count
                On Error Resume Next
                txt = tbl.Cell(r, c).Range.Text
                On Error GoTo 0
                
                txt = Replace(Replace(txt, Chr(13), ""), Chr(7), "")
                
                If InStr(1, txt, "Comments", vbTextCompare) > 0 Then
                    colComments = c
                    rHeader = r
                End If
                If InStr(1, txt, "Status", vbTextCompare) > 0 Then
                    colStatus = c
                    rHeader = r
                End If
                If InStr(1, txt, "Issue", vbTextCompare) > 0 Then
                    colIssue = c
                    rHeader = r
                End If
            Next c
        For Each row_dummy In tbl.Rows: Next ' keep structured context
        Next r
        
        ' Process the row below the headers exactly as requested
        If rHeader > 0 And rHeader + 1 <= tbl.Rows.Count Then
            
            ' --- MODIFICATION & FORCE CENTERING BLOCK INSIDE WORD ---
            On Error Resume Next
            
            ' 1. Process Comments Column
            If colComments > 0 Then
                ' Read existing text to check if it's empty or needs to be "Not passed"
                txt = Replace(Replace(tbl.Cell(rHeader + 1, colComments).Range.Text, Chr(13), ""), Chr(7), "")
                If Trim(txt) = "" And colStatus > 0 Then
                    ' Optional condition matching your FillNP: if it's empty, update it
                    tbl.Cell(rHeader + 1, colComments).Range.Text = "Not passed"
                    txt = "Not passed"
                End If
                currentComment = Trim(txt)
                
                ' ALIGN CENTER IN THE MIDDLE (Horizontal + Vertical Center alignment)
                tbl.Cell(rHeader + 1, colComments).Range.ParagraphFormat.Alignment = 1 ' Center Horizontally
                tbl.Cell(rHeader + 1, colComments).VerticalAlignment = 1               ' Center Vertically (wdCellAlignVerticalCenter)
            End If
            
            ' 2. Process Status Column
            If colStatus > 0 Then
                txt = Replace(Replace(tbl.Cell(rHeader + 1, colStatus).Range.Text, Chr(13), ""), Chr(7), "")
                ' If your macro targets modifying it to "NP"
                If Trim(txt) = "" Or UCase(Trim(txt)) = "NP" Then
                    tbl.Cell(rHeader + 1, colStatus).Range.Text = "NP"
                    txt = "NP"
                End If
                currentStatus = UCase(Trim(txt))
                
                ' ALIGN CENTER IN THE MIDDLE (Horizontal + Vertical Center alignment)
                tbl.Cell(rHeader + 1, colStatus).Range.ParagraphFormat.Alignment = 1 ' Center Horizontally
                tbl.Cell(rHeader + 1, colStatus).VerticalAlignment = 1               ' Center Vertically (wdCellAlignVerticalCenter)
            End If
            
            ' 3. Process Issue Column
            If colIssue > 0 Then
                txt = Replace(Replace(tbl.Cell(rHeader + 1, colIssue).Range.Text, Chr(13), ""), Chr(7), "")
                currentIssue = Trim(txt)
                
                ' ALIGN CENTER IN THE MIDDLE (Horizontal + Vertical Center alignment)
                tbl.Cell(rHeader + 1, colIssue).Range.ParagraphFormat.Alignment = 1 ' Center Horizontally
                tbl.Cell(rHeader + 1, colIssue).VerticalAlignment = 1               ' Center Vertically (wdCellAlignVerticalCenter)
            End If
            On Error GoTo 0
            
            ' --- WRITE EXTRACTED DATA TO EXCEL DASHBOARD ---
            If currentTestID <> "" Then
                Call WriteToExcelFinal(excelWs, currentTestID, currentStatus, currentIssue, currentComment, nokRow, okRow, npRow, emptyRow)
            End If
        End If
    Next tbl
    
    ' ==========================================
    ' 3. POST-PROCESSING EXCEL SORT & EXCEL CENTERING
    ' ==========================================
    If npRow > 13 Then
        Dim sortRange As Object
        Set sortRange = excelWs.Range("H13:K" & (npRow - 1))
        sortRange.Sort Key1:=excelWs.Range("J13"), Order1:=1, Key2:=excelWs.Range("K13"), Order2:=1, Header:=2
    End If
    
    ' Enforce centering inside Excel worksheet cells
    If nokRow > 13 Then excelWs.Range("B13:C" & (nokRow - 1)).HorizontalAlignment = -4108
    If okRow > 13 Then excelWs.Range("F13:F" & (okRow - 1)).HorizontalAlignment = -4108
    If npRow > 13 Then excelWs.Range("I13:K" & (npRow - 1)).HorizontalAlignment = -4108
    If emptyRow > 13 Then excelWs.Range("N13:N" & (emptyRow - 1)).HorizontalAlignment = -4108
    
    excelWs.Columns("A:O").AutoFit
    MsgBox "Done! Word file tables corrected and centered. Excel Dashboard Generated.", vbInformation
End Sub

Sub WriteToExcelFinal(ws As Object, tID As String, stat As String, iss As String, comm As String, ByRef nRow As Long, ByRef oRow As Long, ByRef pRow As Long, ByRef eRow As Long)
    If Trim(stat) = "" Then stat = "EMPTY"

    Select Case UCase(stat)
        Case "NOK"
            ws.Cells(nRow, 1).Value = tID
            ws.Cells(nRow, 2).Value = stat
            ws.Cells(nRow, 3).Value = iss
            ws.Cells(nRow, 2).HorizontalAlignment = -4108
            ws.Cells(nRow, 3).HorizontalAlignment = -4108
            ws.Range(ws.Cells(nRow, 1), ws.Cells(nRow, 3)).Borders.LineStyle = 1
            nRow = nRow + 1
        Case "OK"
            ws.Cells(oRow, 5).Value = tID
            ws.Cells(oRow, 6).Value = stat
            ws.Cells(oRow, 6).HorizontalAlignment = -4108
            ws.Range(ws.Cells(oRow, 5), ws.Cells(oRow, 6)).Borders.LineStyle = 1
            oRow = oRow + 1
        Case "NP"
            ws.Cells(pRow, 8).Value = tID
            ws.Cells(pRow, 9).Value = stat
            ws.Cells(pRow, 10).Value = iss
            ws.Cells(pRow, 11).Value = comm
            ws.Cells(pRow, 9).HorizontalAlignment = -4108
            ws.Cells(pRow, 10).HorizontalAlignment = -4108
            ws.Cells(pRow, 11).HorizontalAlignment = -4108
            ws.Range(ws.Cells(pRow, 8), ws.Cells(pRow, 11)).Borders.LineStyle = 1
            pRow = pRow + 1
        Case "EMPTY", ""
            ws.Cells(eRow, 13).Value = tID
            ws.Cells(eRow, 14).Value = "MISSING"
            ws.Cells(eRow, 14).HorizontalAlignment = -4108
            ws.Range(ws.Cells(eRow, 13), ws.Cells(eRow, 14)).Borders.LineStyle = 1
            eRow = eRow + 1
    End Select
End Sub