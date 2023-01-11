Sub CombineSheetsToActiveWorkbook()
    Dim ws As Worksheet
    Dim lRow As Long
    Dim header As Range
    Dim targetWB As Workbook

    'Open the target Excel file
    Set targetWB = Workbooks.Open("C:\MyFolder\EXCEL1.xlsx")

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
End Sub
