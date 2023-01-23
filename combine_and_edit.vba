Sub CombineSheetsToActiveWorkbook()
    Dim ws As Worksheet
    Dim lRow As Long
    Dim header As Range
    Dim targetWB As Workbook
    Dim i As Long, j As Long

    'Open the target Excel file
    Set targetWB = Workbooks.Open("C:\Users\yduz\Desktop\macros\12-RSU-December.2022.xlsx")

    ' Copy data from each worksheet in the target workbook to the active workbook
    For Each ws In targetWB.Sheets
        If ws.Name = "BN" Or ws.Name = "LH" Or ws.Name = "ED" Or ws.Name = "Shelter & WASH" Or ws.Name = "PR" Or ws.Name = "Inter-Sector" Or ws.Name = "FSA" Or ws.Name = "Health" Then
            lRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            'copy the header row from the first sheet
            If ws.Name = "BN" Then
                'Check if the range "A1:Z1" exist, if not check the usedRange
                On Error Resume Next
                Set header = ws.Range("A1:Z1")
                If header Is Nothing Then
                    Set header = ws.UsedRange.Rows(1)
                End If
                On Error GoTo 0
                ThisWorkbook.Sheets(1).Range("A1:Z1").Value = header.Value
            End If
            'copy the data starting from row 2
            ws.Range("A2:Z" & lRow).Copy ThisWorkbook.Sheets(1).Range("A" & ThisWorkbook.Sheets(1).Cells(ThisWorkbook.Sheets(1).Rows.Count, 1).End(xlUp).Row + 1)
        End If
    Next ws
    ' Close the target Excel file
    targetWB.Close
    ' Delete the columns 23,24,25,26,27 in the active worksheet
    With ThisWorkbook.Sheets(1)
    For i = 1 To .UsedRange.Rows.Count
        If i <> 1 Then
            If IsNumeric(.Cells(i, 17)) Then
                .Cells(i, 17) = CLng(.Cells(i, 17))
            Else
                .Cells(i, 17).Value = Val(.Cells(i, 17))
            End If
        End If
    Next i
End With

    
    On Error Resume Next
    ThisWorkbook.Sheets(1).Columns(23).Delete
    ThisWorkbook.Sheets(1).Columns(24).Delete
    ThisWorkbook.Sheets(1).Columns(25).Delete
    ThisWorkbook.Sheets(1).Columns(26).Delete
    ThisWorkbook.Sheets(1).Columns(27).Delete
    On Error GoTo 0
    
    'Check for errors in the active worksheet and replace them with 0
    With ThisWorkbook.Sheets(1)
        For i = 1 To .UsedRange.Rows.Count
            For j = 1 To .UsedRange.Columns.Count
                If IsError(.Cells(i, j)) Then
                    .Cells(i, j) = ""
                End If
            Next j
        Next i
    End With
    
    'Change column names
    With ThisWorkbook.Sheets(1)
    .Cells(1, 11).Value = "Boys"
    .Cells(1, 12).Value = "Girls"
    .Cells(1, 13).Value = "Men"
    .Cells(1, 14).Value = "Women"
    End With
    
    'Delete texts in the column 17
    With ThisWorkbook.Sheets(1)
        For i = 1 To .UsedRange.Rows.Count
            If IsNumeric(.Cells(i, 17).Value) Then
                Dim spacePos As Integer
                spacePos = InStr(.Cells(i, 17).Value, " ")
                If spacePos > 0 Then
                    .Cells(i, 17).Value = Left(.Cells(i, 17).Value, spacePos - 1)
                End If
            End If
        Next i
    End With
End Sub


