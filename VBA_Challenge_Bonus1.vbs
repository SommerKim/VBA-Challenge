Sub ticker_tracker_summary():

    Dim ticker_letter As String
    Dim beg_year As Double
    Dim end_year As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim total_volume As Double
    Dim total_start As Range
    Dim total_end As Range
    Dim Summary_Sheet As Worksheet
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    Dim ws As Worksheet
    
    'We're going to be using all the worksheets in this workbook.
    For Each ws In ThisWorkbook.Sheets
        
        ' Make sure the workbook has a Summary Sheet at the first tab.
        If Sheets(1).Name <> "Summary Sheet" Then
        Sheets.Add(Before:=Sheets(1)).Name = "Summary Sheet"
        End If
        
        Set Summary_Sheet = Application.Sheets(1)
        
        'Add headers for new summary table on Summary Sheet.
        Range("A1:D1").Interior.ColorIndex = 37
        Cells(1, 1).Value = "Ticker"
        Cells(1, 2).Value = "Yearly Change"
        Cells(1, 3).Value = "Percent Change"
        Cells(1, 4).Value = "Total Stock Volume"
        Cells(1, 1).ColumnWidth = 10
        Range("B1:D1").ColumnWidth = 16.5
    
        'Cycle through each page in the workbook that isn't Summary Sheet.
        If ws.Name <> "Summary Sheet" Then
        
            ' Turn off screen updating, activate the worksheet, Find last row, search rows 2 through last row.
            Application.ScreenUpdating = False
            ws.Activate
            lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
            For row_num = 2 To lastrow
              
                ' Find the end of the section for a single stock by finding the last row before
                ' the ticker letter changes to a new one.
                If Cells(row_num + 1, 1).Value <> Cells(row_num, 1).Value Then
                    
                    ' Collect ticker letter for the Summary Table row on the Summary Sheet.
                    ticker_letter = Cells(row_num, 1).Value
                    Summary_Sheet.Range("A" & Summary_Table_Row).Value = ticker_letter
                    
                    ' If it's the first ticker letter collected, the stock's opening value (beg_year)
                    ' will not yet have a value because that value is set as we move to populate the
                    ' next row in the Summary Table. Therefore, beg_year will be in cells 2,3 and the
                    ' stock's beginning volume (total_start) will be in 2,7
                    If beg_year = 0 Then
                        beg_year = Cells(2, 3).Value
                        Set total_start = Cells(2, 7)
                    End If
                    
                    ' Set stock's ending value (end_year), calculate the total and yearly change, and
                    ' put the results in Summary Table.
                    end_year = Cells(row_num, 6).Value
                    Set total_end = Cells(row_num, 7)
                    yearly_change = end_year - beg_year
                    Summary_Sheet.Range("B" & Summary_Table_Row).Value = yearly_change
                    
                    ' Color code the yearly_change cells with red for a negative value and green for a positive.
                    If yearly_change < 0 Then
                        Summary_Sheet.Range("B" & Summary_Table_Row).Interior.ColorIndex = 3
                    Else:
                        Summary_Sheet.Range("B" & Summary_Table_Row).Interior.ColorIndex = 43
                    End If
                    
                    ' Calculate the percent change and total volume, then place them on the Summary Table.
                    percent_change = (end_year - beg_year) / beg_year
                    Summary_Sheet.Range("C" & Summary_Table_Row).Value = percent_change
                    total_volume = Application.WorksheetFunction.Sum(Range(total_start, total_end))
                    Summary_Sheet.Range("D" & Summary_Table_Row).Value = total_volume
                  
                    ' Move to populate the next row in Summary Table. Set the beginning year stock value and
                    ' the total stock volume value for the next loop.
                    Summary_Table_Row = Summary_Table_Row + 1
                    beg_year = Cells(row_num + 1, 3).Value
                    total_start = Cells(row_num + 1, 7)

                End If
            Next row_num
        End If
    Next ws
    Application.ScreenUpdating = True

End Sub


