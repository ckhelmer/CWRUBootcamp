Attribute VB_Name = "Module4"
'TAs: I separated the assignment out into two separate subs because the easy/moderate assignment runs fine, but the hard assignment completely locks up excel on my machine. It's the second subroutine, YearlyMinMax.

Sub TickerTotal()

'Outer worksheet loop
Dim ws As Worksheet
For Each ws In ActiveWorkbook.Worksheets
ws.Activate
   
   'Declare and assign variables
   Dim Summary_Table_Row As Integer
   Summary_Table_Row = 2
   
   Dim LastRow As Long
   LastRow = Cells(Rows.Count, 1).End(xlUp).Row
   
   Dim TickerName As String
   
   Dim Volume As Double
   Volume = 0
      
    Dim FirstDateOpen As Double
    Dim LastDateClose As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    

'Put headers in the correct place
   Range("I1").Value = "Ticker"
   Range("J1").Value = "Yearly Change"
   Range("K1").Value = "Percent Change"
   Range("L1").Value = "Total Stock Volume"


'Format row K for % change
    Range("K:K").NumberFormat = "0.00%"
   
  
'Inner loop: TickerName, Volume, Yearly Change,

   For i = 2 To LastRow
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
            'Set Ticker Name
            TickerName = Cells(i, 1).Value
            
            'Count Volume
            Volume = Volume + Cells(i, 7).Value
            
            'YearlyChange
                If Cells(i, 2) = WorksheetFunction.Min(Cells(i, 2)) Then
                    FirstDateOpen = Cells(i, 3).Value
                End If
                If Cells(i, 2) = WorksheetFunction.Max(Cells(i, 2)) Then
                    LastDateOpen = Cells(i, 6).Value
                End If
                
                YearlyChange = LastDateOpen - FirstDateOpen
                
                     
             'PercentChange
                If LastDateOpen <> 0 Then
                    PercentChange = YearlyChange / LastDateOpen
                                  
                End If
                
                            
            'Display results
            Range("I" & Summary_Table_Row) = TickerName
            Range("J" & Summary_Table_Row) = YearlyChange
            Range("K" & Summary_Table_Row) = PercentChange
            Range("L" & Summary_Table_Row) = Volume
            
                       
            'Move next entry in summary table down
            Summary_Table_Row = Summary_Table_Row + 1
            
            'Reset Values
            YearlyChange = 0
            Volume = 0
            PercentChange = 0
            
        Else
                        
             'Count Volume
             Volume = Volume + Cells(i, 7).Value
        
                  
        
        End If
              
   Next i
   
   'Second Loop for Conditional Formatting, Max, Min, and Greatest Total Volume
   

   For j = 2 To LastRow
        If Cells(j, 10).Value > 0 Then
            Cells(j, 10).Interior.ColorIndex = 4
        ElseIf Cells(j, 10).Value <= 0 And Cells(j, 10).Value <> " " Then
            Cells(j, 10).Interior.ColorIndex = 3
        End If
        
  Next j

    'Reformat Columns:
    Columns("A:N").AutoFit
Next ws
End Sub

Sub YearlyMinMax()
'This is the sub for the hard assignment. It works on the test sheet, but it locks up excel on the yearly document.

Dim ws As Worksheet
For Each ws In ActiveWorkbook.Worksheets
ws.Activate

    Dim GreatestIncrease As Double
    Dim GreatestDecrease As Double
    Dim GreatestVolume As Double
    
    Dim LastRow As Long
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'Summary Table Headings
    Range("N2").Value = "Greatest Percent Increase"
    Range("N3").Value = "Greatest Percent Decrease"
    Range("N4").Value = "Greatest Total Volume"
    Range("O1").Value = "Ticker"
    Range("P1").Value = "Value"
    
    For k = 2 To LastRow
    

    'Greatest % Increase
        If Cells(k, 11) = WorksheetFunction.Max(Range("K:K")) Then
            GreatestIncrease = Cells(k, 11).Value
            Range("O2").Value = Cells(k, 9).Value
            Range("P2").Value = GreatestIncrease
            
        End If
   'Greatest % Decrease
  
        If Cells(k, 11) = WorksheetFunction.Min(Range("K:K")) Then
            GreatestDecrease = Cells(k, 11).Value
            Range("O3").Value = Cells(k, 9).Value
            Range("P3").Value = GreatestDecrease
        End If
   'Greatest Total Volume
   
        If Cells(k, 12) = WorksheetFunction.Max(Range("L:L")) Then
            GreatestVolume = Cells(k, 12).Value
            Range("O4").Value = Cells(k, 9).Value
            Range("P4").Value = GreatestVolume
        End If
    
        
   Next k
   
   'Reformat rows for %
    Range("P2:P3").NumberFormat = "0.00%"
    
    'Reformat rows for width
    Columns("N:P").AutoFit

Next ws
End Sub

