Sub stocks_loop():

' Create a script that will loop through all the stocks for one year and output the following information.
' Set an initial variable for holding the ticker symbol, opening price, closing price, and total stock volume.
' capture ticker
Dim ticker_symbol As Long
Dim opening_price As Long
Dim closing_price As Long

' continiously add to volume throughout year - ex: G2 + G3 + G4, etc.

Dim total_stock_volume As Long
total_stock_volume = 0

' output for the iterations we are trying to find
Dim final_ticker As Long
Dim yearly_change As Long
Dim percent_change As Long

' this will be used as the summary row - intial row, do not forget to add 1 as we go down.
Dim summary_row As Integer
summary_row = 2

' Make labels for boxes with titles Via Ranges

    ' Insert Data Titles Via Ranges by Declaing Strings
    
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change (%)"
    Range("L1").Value = "Total Stock Volume"

' Loop through all stocks for row 2-70926

For i = 2 To 70926
   

' at every i check current cell + 1
' (output * The ticker symbol)
' i = row in this case, while 1 is the column. cells(row, col)
' cells(i, 1).Value ()
' assign ticker_symbol a value, in this case a string by using cells(i, j).Value

ticker_symbol = Cells(i, 1).Value
opening_price = Cells(i, 3).Value

If ticker_symbol <> Cells(i + 1, 1).Value Then

    ' capture our closing price
    
    closing_price = Cells(i, 6).Value

    ' we need to calculate yearly change (closing_price - opening_price)
   
    yearly_change = (closing_price - opening_price)

    ' we need to calculate percent change
   
    percent_change = (closing_price - opening_price / opening_price)

    ' we need to also calculate total volume
   
    total_stock_volume = total_stock_volume + Cells(i, 7).Value

    ' set summary table values
    ' "we are saying go to this cell and assign it a value of the ticker"
     
     Cells(summary_row, 9).Value = ticker_symbol

     Cells(summary_row, 10).Value = yearly_change

     Cells(summary_row, 11).Value = percent_change

     Cells(summary_row, 12).Value = total_stock_volume


    ' we are going to reset values and iterate counters so that we have a fresh set for our next iteration
    
    summary_row = summary_row + 1
    closing_price = 0
    yearly_change = 0
    percent_change = 0
    total_stock_volume = 0


Else

    ' we need to keep tabs on a running sum of the stock volume until the end of the stock we are currently iterating
    
    total_stock_volume = total_stock_volume + Cells(i, 7).Value


End If

Next i 

End Sub



