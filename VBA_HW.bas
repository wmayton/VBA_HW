Attribute VB_Name = "Module1"
Sub stock_prices()

Dim Ticker_Symbol as String.??Dim Stock_Price as Double

Dim Row_Count As Long


    For i = 2 To Row_Count


        ' Check if we are still within the ticker symbol, if it is not...
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

                ' Set the Ticker Symbol
                Ticker_Symbol = Cells(i, 1).Value

      ' Add to the Stock Price Total
      Stock_Total = Stock_Total + Cells(i, 5).Value

      ' Print the Ticker Symbol in the Summary Table
      Range("G" & Summary_Table_Row).Value = Ticker_Symbol

      ' Print the Stock Total to the Summary Table
      Range("H" & Summary_Table_Row).Value = Stock_Total

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Stock Total
      Stock_Total = 0

    ' If the cell immediately following a row is the same brand...
    Else



PSEUDOCODE: Print
Set a variable for ticker name?Set a variable for row count?Set a variable for earliest value, opening price?Set a variable for stock volume?
For every row:?

    IF  ?       Ticker symbol in row we are on= initial ticker symbol, then keep going?     Subtract opening price in the first row from closing price in current row
        Add value of 7th column in each row [to get the total volume].
            Print value in column 4.??

    ELSE?       Print the difference of closing value in current row and the opening price of the first row of      given ticker symbol.
        If value is positive, green. If value is negative, red).?       Convert that difference into percentage. Print percentage value in third column.
        Move down one row, assign new ticker symbol value.
?   ?   Then, print this value expressed as a percentage in third column of my table.?? Then, I want it to find sum of each volume for any given ticker symbol, and print this is column 4  of my table.
