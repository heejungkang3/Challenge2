Attribute VB_Name = "Module2"
Sub Title()
    For Each ws In Worksheets

        Range("I1") = "Ticker"
        Range("J1") = "Yearly Change"
        Range("K1") = "Percent Change"
        Range("L1") = "Total Stock Volume"
    Next ws

   
End Sub
Sub Ticker_TotalVol()

    For Each ws In Worksheets
        Dim Ticker As String
        Dim i As Integer
        Dim Lastrow As Long
        Dim TickerRow As Integer
        Dim TotalVol As Double
        
        TotalVol = 0
       
        TickerRow = 2
               
        Lastrow = Cells(Rows.Count, 1).End(xlUp).Row
       
        For i = 2 To 30000
           
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
           
                Ticker = Cells(i, 1).Value
               
                TotalVol = TotalVol + Cells(i, 7).Value
               
                Range("I" & TickerRow).Value = Ticker
               
                Range("L" & TickerRow).Value = TotalVol
               
                TickerRow = TickerRow + 1
               
                TotalVol = 0
               
            Else
           
                TotalVol = TotalVol + Cells(i, 7).Value
               
            End If
           
        Next i
       
    'VBA could not run all in one for loop as the lastrow was too big of an integer
        For j = 30001 To Lastrow
          If Cells(j + 1, 1).Value <> Cells(j, 1).Value Then
           
                Ticker = Cells(j, 1).Value
               
                TotalVol = TotalVol + Cells(j, 7).Value
               
                Range("I" & TickerRow).Value = Ticker
               
                Range("L" & TickerRow).Value = TotalVol
               
                TickerRow = TickerRow + 1
               
                TotalVol = 0
               
            Else
           
                TotalVol = TotalVol + Cells(j, 7).Value
               
            End If
        Next j
   Next ws
End Sub

Sub YearlyChange()
    For Each ws In Worksheets
        Dim i As Integer
        Dim YearlyChange As String
        Dim ChangeRow As Integer
        Dim Lastrow As Long
        Dim YearClose As Double
        Dim YearOpen As Double
        Dim PercentChange As Double
           
        ChangeRow = 2
       
        Lastrow = Cells(Rows.Count, 1).End(xlUp).Row
        
        YearOpen = Range("C2").Value
       
        For i = 2 To 20000
    
            Range("M" & ChangeRow).Value = YearOpen
               
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
               
                YearClose = Cells(i, 6).Value
               
                YearlyChange = YearClose - YearOpen
               
                PercentChange = (YearClose - YearOpen) / YearOpen
                           
                Range("J" & ChangeRow).Value = YearlyChange
                                       
                Range("K" & ChangeRow).Value = PercentChange
                Range("K" & ChangeRow).NumberFormat = "0.00%"
               
                ChangeRow = ChangeRow + 1
                
                YearOpen = Cells(i, 3).Value
               
            End If
             
            If Cells(i, 10) < 0 Then
               
                Cells(i, 10).Interior.ColorIndex = 3
           
            Else
                   
                Cells(i, 10).Interior.ColorIndex = 4
            End If
         Next i
         
         For j = 20001 To Lastrow
    
            Range("M" & ChangeRow).Value = YearOpen
               
            If Cells(j + 1, 1).Value <> Cells(j, 1).Value Then
               
                YearClose = Cells(j, 6).Value
               
                YearlyChange = YearClose - YearOpen
               
                PercentChange = (YearClose - YearOpen) / YearOpen
                           
                Range("J" & ChangeRow).Value = YearlyChange
                                       
                Range("K" & ChangeRow).Value = PercentChange
                Range("K" & ChangeRow).NumberFormat = "0.00%"
               
                ChangeRow = ChangeRow + 1
                
                YearOpen = Cells(j, 3).Value
               
            End If
             
            If Cells(j, 10) < 0 Then
               
                Cells(j, 10).Interior.ColorIndex = 3
           
            Else
                   
                Cells(j, 10).Interior.ColorIndex = 4
            End If
         Next j
   Next ws
End Sub

Sub ExtremeChange()
  For Each ws In Worksheets
    Range("P2").Value = "Greatest % Increase"
    Range("P3").Value = "Greatest % Decrease"
    Range("P4").Value = "Greatest Total Volume"
    Range("Q1").Value = "Ticker"
    Range("R1").Value = "Value"
    
    
   
    
   
  Next ws
End Sub
