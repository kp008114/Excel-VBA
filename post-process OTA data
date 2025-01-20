Sub ota()

    Dim lastRow As Integer
    With ActiveSheet
    lastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
    End With
    
    Application.DisplayAlerts = False
    
    Dim string_curr As String
    Dim arrSplitStrings1() As String
    Dim cal_n As Integer
    Dim cal_s As String
    FCAL_n = 0   'FCAL次數
    'cal_s = "RF Cal (RF1):"
    
    Dim NewBook As Workbook
    Set NewBook = Workbooks.Add
    With NewBook
        .Worksheets(1).Name = "計算"
        .Sheets.Add After:=Sheets(1)
        .Sheets(2).Name = "總表"
        .Sheets(2).Cells(1, 1).Value = "Rx1"
        .Sheets(2).Cells(1, 5).Value = "Rx2"
        .Sheets(2).Cells(2, 1).Value = "Idx"
        .Sheets(2).Cells(2, 2).Value = "S"
        .Sheets(2).Cells(2, 3).Value = "N"
        .Sheets(2).Cells(2, 4).Value = "SNR"
        .Sheets(2).Cells(2, 5).Value = "Idx"
        .Sheets(2).Cells(2, 6).Value = "S"
        .Sheets(2).Cells(2, 7).Value = "N"
        .Sheets(2).Cells(2, 8).Value = "SNR"
    End With
    
    Dim form_row_curr As Integer '總表的目前列數
    form_row_curr = 3
    
    'Dim sheet As Worksheet
    'Set sheet = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
    
    Workbooks("OTA_cal.xlsm").Activate
    
    For row_curr = 1 To lastRow
    '================================FCAL data=================================
        If Cells(row_curr, 1).Value = "RF Cal (RF1):" Then

            FCAL_n = FCAL_n + 1
            
            row_curr = row_curr + 6
        
        '================================Rx1 and Rx2 data=================================
        ElseIf Cells(row_curr, 1).Value = "FDECT Rx1:" Then
        
        '================================Rx1=================================
            string_curr = Cells(row_curr + 1, 1).Value
            Do
                temp = string_curr
                string_curr = Replace(string_curr, "  ", " ") 'remove multiple white spaces
            Loop Until temp = string_curr
            arrSplitStrings1 = Split(string_curr, " ")
            NewBook.Sheets("總表").Cells(form_row_curr, 1).Value = arrSplitStrings1(2)
            
            string_curr = Cells(row_curr + 2, 1).Value
            Do
                temp = string_curr
                string_curr = Replace(string_curr, "  ", " ") 'remove multiple white spaces
            Loop Until temp = string_curr
            'Debug.Print string_cur
            arrSplitStrings1 = Split(string_curr, " ")
            NewBook.Sheets("總表").Cells(form_row_curr, 2).Value = arrSplitStrings1(2)
            
            string_curr = Cells(row_curr + 3, 1).Value
            Do
                temp = string_curr
                string_curr = Replace(string_curr, "  ", " ") 'remove multiple white spaces
            Loop Until temp = string_curr
            'Debug.Print string_cur
            arrSplitStrings1 = Split(string_curr, " ")
            NewBook.Sheets("總表").Cells(form_row_curr, 3).Value = arrSplitStrings1(2)
            
            string_curr = Cells(row_curr + 4, 1).Value
            Do
                temp = string_curr
                string_curr = Replace(string_curr, "  ", " ") 'remove multiple white spaces
            Loop Until temp = string_curr
            'Debug.Print string_cur
            arrSplitStrings1 = Split(string_curr, " ")
            NewBook.Sheets("總表").Cells(form_row_curr, 4).Value = arrSplitStrings1(2)
            
            
            
    '        ElseIf Cells(row_curr, 1).Value = "FDECT Rx2:" Then
    '
    '        row_curr = row_curr + 5
    
            '================================Rx2=================================
            string_curr = Cells(row_curr + 7, 1).Value
            Do
                temp = string_curr
                string_curr = Replace(string_curr, "  ", " ") 'remove multiple white spaces
            Loop Until temp = string_curr
            arrSplitStrings1 = Split(string_curr, " ")
            NewBook.Sheets("總表").Cells(form_row_curr, 5).Value = arrSplitStrings1(2)
            
            string_curr = Cells(row_curr + 8, 1).Value
            Do
                temp = string_curr
                string_curr = Replace(string_curr, "  ", " ") 'remove multiple white spaces
            Loop Until temp = string_curr
            arrSplitStrings1 = Split(string_curr, " ")
            NewBook.Sheets("總表").Cells(form_row_curr, 6).Value = arrSplitStrings1(2)
            
            string_curr = Cells(row_curr + 9, 1).Value
            Do
                temp = string_curr
                string_curr = Replace(string_curr, "  ", " ") 'remove multiple white spaces
            Loop Until temp = string_curr
            arrSplitStrings1 = Split(string_curr, " ")
            NewBook.Sheets("總表").Cells(form_row_curr, 7).Value = arrSplitStrings1(2)
            
            string_curr = Cells(row_curr + 10, 1).Value
            Do
                temp = string_curr
                string_curr = Replace(string_curr, "  ", " ") 'remove multiple white spaces
            Loop Until temp = string_curr
            arrSplitStrings1 = Split(string_curr, " ")
            NewBook.Sheets("總表").Cells(form_row_curr, 8).Value = arrSplitStrings1(2)
            
            form_row_curr = form_row_curr + 1
            row_curr = row_curr + 11
        
        End If
        
    Next
    
    Sheets("工作表1").UsedRange.ClearContents
    
    string_curr = Application.ActiveWorkbook.Path
    NewBook.Close SaveChanges:=True, Filename:=string_curr & "\" & Format(DateTime.Now, "yyyyMMdd_hhmmss")

    Debug.Print "hello"
End Sub


