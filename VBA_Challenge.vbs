Sub ticker_tracker():

    Dim ticker_letter As String
    Dim beg_year As Double
    Dim end_year As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim total_volume As Double
    Dim Summary_Sheet As Worksheet
    Set Summary_Sheet = Application.Sheets(1)
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    Dim ws As Worksheet
    
    
    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> "Summary Sheet" Then
          ws.Activate
          lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
          For row_num = 2 To lastrow
              
                  If Cells(row_num + 1, 1).Value <> Cells(row_num, 1).Value Then
                  ticker_letter = Cells(row_num, 1).Value
                  Summary_Sheet.Range("A" & Summary_Table_Row).Value = ticker_letter
                    If beg_year = 0 Then
                    beg_year = Cells(2, 3).Value
                    End If
                  end_year = Cells(row_num, 6)
                  yearly_change = end_year - beg_year
                  Summary_Sheet.Range("B" & Summary_Table_Row).Value = yearly_change
                    If yearly_change < 0 Then
                    Summary_Sheet.Range("B" & Summary_Table_Row).Interior.ColorIndex = 3
                    Else: Summary_Sheet.Range("B" & Summary_Table_Row).Interior.ColorIndex = 43
                    End If
                  percent_change = (end_year - beg_year) / beg_year
                  Summary_Sheet.Range("C" & Summary_Table_Row).Value = percent_change
                  Summary_Table_Row = Summary_Table_Row + 1
                  beg_year = Cells(row_num + 1, 3).Value
                  
              End If
          Next row_num
        End If
    Next ws

End Sub