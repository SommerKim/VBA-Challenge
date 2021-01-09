Sub greatest():

    Dim great_increase_ticker As String
    Dim great_increase_value As Double
    Dim great_decrease_ticker As String
    Dim great_decrease_value As Double
    Dim great_volume_ticker As String
    Dim great_volume_value As Double
    Dim great_increase_match As Double
    Dim great_decrease_match As Double
    Dim great_volume_match As Double
    Dim percent_range As Range
    Dim total_range As Range
    Dim table_range As Range
    
    ' Find last row
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
      
    ' We will use these ranges to find the stock with the greatest percent increase (great_increase_value and
    ' great_increase_ticker), the greatest percent decrease (great_decrease_value and great_decrease_ticker)
    ' and also the max total volume (great_volume_value and great_volume_ticker).
    Set percent_range = Range("C2:C" & lastrow)
    Set table_range = Range("A2:D" & lastrow)
    
    ' Find the max value in the percent change column, then use .Match and .Index to find its
    ' corresponding ticker.
    great_increase_value = Application.WorksheetFunction.Max(percent_range)
    great_increase_match = WorksheetFunction.Match(great_increase_value, percent_range, 0)
    great_increase_ticker = WorksheetFunction.Index(table_range, great_increase_match, 1)
    Cells(2, 7).Value = great_increase_ticker
    Cells(2, 8).Value = great_increase_value
        
    ' Find the min value in the percent change column, then use .Match and .Index to find its
    ' corresponding ticker.
    great_decrease_value = Application.WorksheetFunction.Min(percent_range)
    great_decrease_match = WorksheetFunction.Match(great_decrease_value, percent_range, 0)
    great_decrease_ticker = WorksheetFunction.Index(table_range, great_decrease_match, 1)
    Cells(3, 7).Value = great_decrease_ticker
    Cells(3, 8).Value = great_decrease_value
    
    ' Find the max value in the total stock volume column, then use .Match and .Index to find
    ' its corresponding ticker.
    Set total_range = Range("D2:D" & lastrow)
    great_volume_value = Application.WorksheetFunction.Max(total_range)
    great_volume_match = WorksheetFunction.Match(great_volume_value, total_range, 0)
    great_volume_ticker = WorksheetFunction.Index(table_range, great_volume_match, 1)
    Cells(4, 7).Value = great_volume_ticker
    Cells(4, 8).Value = great_volume_value
    
    
End Sub
