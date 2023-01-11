Sub CombineSheets()
    Dim ws As Worksheet
    Dim newWs As Worksheet
    Dim lRow As Long
    Dim header As Range

    ' Add a new worksheet to the workbook
    Set newWs = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    newWs.Name = "Combined Data6"

    ' Copy data from each worksheet to the new worksheet
    For Each ws In ThisWorkbook.Sheets
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
                newWs.Range("A1:Z1").Value = header.Value
            End If
            ws.Range("A2:Z" & lRow).Copy newWs.Range("A" & newWs.Cells(newWs.Rows.Count, 1).End(xlUp).Row + 1)
        End If
    Next ws
End Sub
