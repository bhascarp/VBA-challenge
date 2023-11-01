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
