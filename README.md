# VBA-challenge
Please find attached the source code file, to view the entire code. Everything went well, except below two problems.

Two glitches, which I could not resolve in the end:

1. For the yearly change (close value - open value), I used below code, and for some reason the open value is returning as zero and hence the change is not reflecting correctly.
   Hence used the first line of code, that is not commented below. ' Appreciate if you can show me what's wrong here in the commented code.

  ' Set the open value of each ticker
            open_value = Cells(2, 3).Value
            
'            If ws.Name = "2018" Then
'                begin_year_date = "20180102"
'            ElseIf ws.Name = "2019" Then
'                begin_year_date = "20190102"
'            ElseIf ws.Name = "2020" Then
'                begin_year_date = "20200102"
'            End If
'
'            If Cells(i, 2).Value = begin_year_date Then
'                open_value = Cells(i, 3).Value
'            End If


2. I could not show the ticker in the final results on the right most area of the excel file. Below code did not work. 

 'Below code to show the ticker on the right most summary did not work.
  ' Appreciate if you can show me what's wrong here
  
' To set the ticker value in final summary
        topper = Range("R2")
        bottom = Range("R3")
        top_volume = Range("R4")
  
        If Cells(i, 12).Value = topper Then
          ticker_max = Cells(i, 10).Value
        End If
  
        If Cells(i, 12).Value = bottom Then
          ticker_min = Cells(i, 10).Value
        End If
  
        If Cells(i, 13).Value = top_volume Then
          ticker_max_volume = Cells(i, 10).Value
        End If
  
        Range("Q2").Value = ticker_max
        Range("Q3").Value = ticker_min
        Range("Q4").Value = ticker_max_volume

====================================================================================================================================================================================
Below is the full code
====================================================================================================================================================================================

Sub Multi_Year_Stock()

  ' Define the worksheet
    Dim ws As Worksheet

  ' Loop through each worksheet
    For Each ws In Worksheets
    
    ' Activate the worksheet
    ws.Activate

  ' Set a variable to hold the ticker
        Dim Ticker As String
  
  ' Set an initial variable to hold the total stock volume per ticker
        Dim TotalStockVolume As Double
  
  ' To list each ticker and related attributes in summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
  
  ' Set yearly change attribute for each ticker
        Dim Yearly_Change As Double
       
  ' Set percent change of yearly change for each ticker
        Dim Percent_Change As Double
  
  ' Define begin and end dates of the year
        Dim begin_year_date As Long
        Dim end_year_date As Long
  
  ' Assign the dates from the data to the variables
        'begin_year_date = 20180102
        'end_year_date = 20181231
  
  ' Set the open and close values for each ticker
        Dim open_value As Double
        Dim close_value As Double
  
 ' Counts the number of rows
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row

  ' Loop through all stock listings
        For i = 2 To lastrow
  
    ' Check if it is same ticker, if not...
          If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      'Set the ticker name
            Ticker = Cells(i, 1).Value
      
      ' Add to the total stock volume
            TotalStockVolume = TotalStockVolume + Cells(i, 7).Value
      
      
      
      ' Set the open value of each ticker
            open_value = Cells(2, 3).Value
            
'            If ws.Name = "2018" Then
'                begin_year_date = "20180102"
'            ElseIf ws.Name = "2019" Then
'                begin_year_date = "20190102"
'            ElseIf ws.Name = "2020" Then
'                begin_year_date = "20200102"
'            End If
'
'            If Cells(i, 2).Value = begin_year_date Then
'                open_value = Cells(i, 3).Value
'            End If
                                         
      
      ' Set the closing value of each ticker
            If ws.Name = "2018" Then
                end_year_date = "20181231"
            ElseIf ws.Name = "2019" Then
                end_year_date = "20191231"
            ElseIf ws.Name = "2020" Then
                end_year_date = "20201231"
            End If
      
            If Cells(i, 2).Value = end_year_date Then
               close_value = Cells(i, 6).Value
            End If
      
      ' Calculate the yearly change value
            Yearly_Change = close_value - open_value
      
      ' Calculate the Percent Change value
            Percent_Change = Yearly_Change / open_value
      
      
      ' Print all the above stock attributes in the summary table for each ticker
            Range("J" & Summary_Table_Row).Value = Ticker
            Range("M" & Summary_Table_Row).Value = TotalStockVolume
            Range("K" & Summary_Table_Row).Value = Yearly_Change
            Range("L" & Summary_Table_Row).Value = Percent_Change
            Range("L" & Summary_Table_Row).NumberFormat = "0.00%"
      
      ' Format the changes with colors: Green for positive and Red for negative
            If Range("K" & Summary_Table_Row).Value < 0 Then
              Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
            Else
              Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
        
        
            End If
      
      ' Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
      
    
    ' If the cell immediately following a row is the same ticker
          Else
    
      ' Add to the stock total volume
            TotalStockVolume = TotalStockVolume + Cells(i, 7).Value

          End If
    
        Next i
  
  ' High level summary table to show the greatest values
  
  ' Get the Greatest % increase
        Range("R2") = WorksheetFunction.Max(Range("L:L"))
        Range("R2").NumberFormat = "0.00%"
  
  ' Get the Greatest % decrease
        Range("R3") = WorksheetFunction.Min(Range("L:L"))
        Range("R3").NumberFormat = "0.00%"
  
  ' Get the Greatest total volume
        Range("R4") = WorksheetFunction.Max(Range("M:M"))
  
  'Below code to show the ticker on the right most summary did not work.
  ' Appreciate if you can show me what's wrong
  
' To set the ticker value in final summary
        topper = Range("R2")
        bottom = Range("R3")
        top_volume = Range("R4")
  
        If Cells(i, 12).Value = topper Then
          ticker_max = Cells(i, 10).Value
        End If
  
        If Cells(i, 12).Value = bottom Then
          ticker_min = Cells(i, 10).Value
        End If
  
        If Cells(i, 13).Value = top_volume Then
          ticker_max_volume = Cells(i, 10).Value
        End If
  
        Range("Q2").Value = ticker_max
        Range("Q3").Value = ticker_min
        Range("Q4").Value = ticker_max_volume
  
  ' Set the column names and value titles
        Range("J1").Value = "Ticker"
        Range("K1").Value = "Yearly Change"
        Range("L1").Value = "Percent Change"
        Range("M1").Value = "Total Stock Volume"
        Range("Q1").Value = "Ticker"
        Range("R1").Value = "Value"
        Range("P2").Value = "Greatest % Increase"
        Range("P3").Value = "Greatest % Decrease"
        Range("P4").Value = "Greatest Total Volume"
  
    Next ws
    
    MsgBox ("Done")
    
End Sub


================================================================================================================================================================================

        
