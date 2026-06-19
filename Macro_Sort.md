Sub ExtractTestDataWithNP_CleanCommentsAndTitles()
    Dim wordDoc As Document
    Dim excelApp As Object
    Dim excelWb As Object
    Dim excelWs As Object
    Dim regExID As Object, regExIssue As Object
    Dim matches As Object
    Dim para As Paragraph
    Dim currentText As String
    
    Dim currentTestID As String
    Dim currentStatus As String
    Dim currentIssue As String
    Dim currentComment As String
    Dim collectingComments As Boolean
    
    Dim nokRow As Long, okRow As Long, npRow As Long, emptyRow As Long
    
    Set wordDoc = ActiveDocument
    
    ' Setup Regex
    Set regExID = CreateObject("VBScript.RegExp")
    regExID.Pattern = "SWTP_[A-Z0-9_]+"
    regExID.IgnoreCase = True
    
    Set regExIssue = CreateObject("VBScript.RegExp")
    regExIssue.Pattern = "CYBSAF-\d+"
    regExIssue.IgnoreCase = True
    
    ' Initialize Excel
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
    ' 1. STATISTICS / DASHBOARD SECTION (Top Left)
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
    
    ' ==========================================
    ' 2. INDIVIDUAL TABLES SETUP
    ' ==========================================
    Dim startRow As Long
    startRow = 11
    
    ' Table 1: NOK
    excelWs.Cells(startRow, 1).Value = "Failed Tests (NOK)"
    excelWs.Cells(startRow, 1).Font.Bold = True
    excelWs.Cells(startRow + 1, 1).Value = "Test ID"
    excelWs.Cells(startRow + 1, 2).Value = "Status"
    excelWs.Cells(startRow + 1, 3).Value = "Associated Issue"
    excelWs.Range("A12:C12").Interior.Color = RGB(140, 45, 25)
    excelWs.Range("A12:C12").Font.Color = RGB(255, 255, 255)
    excelWs.Range("A12:C12").Font.Bold = True
    nokRow = startRow + 2
    
    ' Table 2: OK
    excelWs.Cells(startRow, 5).Value = "Passed Tests (OK)"
    excelWs.Cells(startRow, 5).Font.Bold = True
    excelWs.Cells(startRow + 1, 5).Value = "Test ID"
    excelWs.Cells(startRow + 1, 6).Value = "Status"
    excelWs.Range("E12:F12").Interior.Color = RGB(46, 105, 48)
    excelWs.Range("E12:F12").Font.Color = RGB(255, 255, 255)
    excelWs.Range("E12:F12").Font.Bold = True
    okRow = startRow + 2
    
    ' Table 3: NP
    excelWs.Cells(startRow, 8).Value = "Not Processed (NP)"
    excelWs.Cells(startRow, 8).Font.Bold = True
    excelWs.Cells(startRow + 1, 8).Value = "Test ID"
    excelWs.Cells(startRow + 1, 9).Value = "Status"
    excelWs.Cells(startRow + 1, 10).Value = "Associated Issue"
    excelWs.Cells(startRow + 1, 11).Value = "Comments (No Issue)"
    excelWs.Range("H12:K12").Interior.Color = RGB(74, 101, 114)
    excelWs.Range("H12:K12").Font.Color = RGB(255, 255, 255)
    excelWs.Range("H12:K12").Font.Bold = True
    npRow = startRow + 2
    
    ' Table 4: EMPTY STATUS
    excelWs.Cells(startRow, 13).Value = "Missing Status (EMPTY)"
    excelWs.Cells(startRow, 13).Font.Bold = True
    excelWs.Cells(startRow + 1, 13).Value = "Test ID"
    excelWs.Cells(startRow + 1, 14).Value = "Status"
    excelWs.Range("M12:N12").Interior.Color = RGB(218, 165, 32)
    excelWs.Range("M12:N12").Font.Color = RGB(255, 255, 255)
    excelWs.Range("M12:N12").Font.Bold = True
    emptyRow = startRow + 2

    ' ==========================================
    ' 3. PARSING DATA FROM WORD
    ' ==========================================
    For Each para In wordDoc.Paragraphs
        currentText = Trim(Replace(para.Range.text, Chr(13), ""))
        currentText = Replace(currentText, Chr(7), "")
        
        ' Rule 1: Stop collecting immediately if we exit a table layout structure
        If para.Range.Tables.Count = 0 Then
            collectingComments = False
        End If
        
        If regExID.Test(currentText) Then
            If currentTestID <> "" Then
                Call WriteToExcelFinal(excelWs, currentTestID, currentStatus, currentIssue, currentComment, nokRow, okRow, npRow, emptyRow)
            End If
            
            Set matches = regExID.Execute(currentText)
            currentTestID = matches(0).Value
            currentStatus = ""
            currentIssue = ""
            currentComment = ""
            collectingComments = False
        End If
        
        If currentTestID <> "" Then
            Dim testCheck As String
            testCheck = UCase(Trim(currentText))
            
            ' Status handling
            If testCheck = "NOK" Or InStr(1, currentText, ",NOK,", vbTextCompare) > 0 Then
                currentStatus = "NOK"
            ElseIf testCheck = "OK" Or InStr(1, currentText, ",OK,", vbTextCompare) > 0 Then
                currentStatus = "OK"
            ElseIf testCheck = "NP" Or InStr(1, currentText, ",NP,", vbTextCompare) > 0 Then
                currentStatus = "NP"
            End If
            
            ' Issue checking
            If regExIssue.Test(currentText) Then
                Set matches = regExIssue.Execute(currentText)
                currentIssue = matches(0).Value
            End If
            
            ' Comment field harvesting validation rules
            If InStr(1, currentText, "Comments", vbTextCompare) > 0 Then
                collectingComments = True
            ElseIf collectingComments And currentText <> "" And Not regExIssue.Test(currentText) And testCheck <> "NP" And testCheck <> "NOK" And testCheck <> "OK" Then
                
                ' Filter out structural headers and tags
                If InStr(1, testCheck, "STATUS") = 0 And InStr(1, testCheck, "ISSUE") = 0 And InStr(1, testCheck, "[END]") = 0 Then
                    
                    ' Rule 2: Ensure we don't grab numbered subsection headers (e.g., "3.5.3.2")
                    If Not (IsNumeric(Left(currentText, 1)) And InStr(1, currentText, ".") > 0) Then
                        If currentComment = "" Then
                            currentComment = currentText
                        Else
                            currentComment = currentComment & " | " & currentText
                        End If
                    Else
                        ' If it looks like a document section title, break compilation immediately
                        collectingComments = False
                    End If
                End If
                
            End If
        End If
    Next para
    
    ' Final loop flush
    If currentTestID <> "" Then
        Call WriteToExcelFinal(excelWs, currentTestID, currentStatus, currentIssue, currentComment, nokRow, okRow, npRow, emptyRow)
    End If
    
    ' Post-processing sort
    If npRow > 13 Then
        Dim sortRange As Object
        Set sortRange = excelWs.Range("H13:K" & (npRow - 1))
        sortRange.Sort Key1:=excelWs.Range("J13"), Order1:=1, Key2:=excelWs.Range("K13"), Order2:=1, Header:=2
    End If
    
    excelWs.Columns("A:O").AutoFit
    MsgBox "Analysis Complete! External section titles successfully excluded.", vbInformation
End Sub

Sub WriteToExcelFinal(ws As Object, tID As String, stat As String, iss As String, comm As String, ByRef nRow As Long, ByRef oRow As Long, ByRef pRow As Long, ByRef eRow As Long)
    If Trim(stat) = "" Then stat = "EMPTY"

    Select Case UCase(stat)
        Case "NOK"
            ws.Cells(nRow, 1).Value = tID
            ws.Cells(nRow, 2).Value = stat
            ws.Cells(nRow, 3).Value = iss
            ws.Range(ws.Cells(nRow, 1), ws.Cells(nRow, 3)).Borders.LineStyle = 1
            nRow = nRow + 1
        Case "OK"
            ws.Cells(oRow, 5).Value = tID
            ws.Cells(oRow, 6).Value = stat
            ws.Range(ws.Cells(oRow, 5), ws.Cells(oRow, 6)).Borders.LineStyle = 1
            oRow = oRow + 1
        Case "NP"
            ws.Cells(pRow, 8).Value = tID
            ws.Cells(pRow, 9).Value = stat
            
            ' Only print comment text if no Issue field exists
            If Trim(iss) <> "" Then
                ws.Cells(pRow, 10).Value = iss
                ws.Cells(pRow, 11).Value = ""
            Else
                ws.Cells(pRow, 10).Value = ""
                ws.Cells(pRow, 11).Value = comm
            End If
            
            ws.Range(ws.Cells(pRow, 8), ws.Cells(pRow, 11)).Borders.LineStyle = 1
            pRow = pRow + 1
        Case "EMPTY"
            ws.Cells(eRow, 13).Value = tID
            ws.Cells(eRow, 14).Value = "MISSING"
            ws.Range(ws.Cells(eRow, 13), ws.Cells(eRow, 14)).Borders.LineStyle = 1
            eRow = eRow + 1
    End Select
End Sub
