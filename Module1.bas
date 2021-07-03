Attribute VB_Name = "Module1"
'* Create a script that will loop through all the stocks for one year and output the following information:

  '* The ticker symbol.

  '* Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

  '* The percent change from opening price at the beginning of a given year to the closing price at the end of that year.

  '* The total stock volume of the stock.

'* You should also have conditional formatting that will highlight positive change in green and negative change in red.


'* Your solution will also be able to return the stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume"

'* Make the appropriate adjustments to your VBA script that will allow it to run on every worksheet, i.e., every year, just by running the VBA script once.

'* Use the sheet `alphabetical_testing.xlsx` while developing your code. This data set is smaller and will allow you to test faster. Your code should run on this file in less than 3-5 minutes.

'* Make sure that the script acts the same on every sheet. The joy of VBA is that it takes the tediousness out of repetitive tasks with a click of the button.

Sub Button1_Click()

' - Set the Variables that need to be kept / used
' - total keeps the Total Stock Volume - i is the start of the current iteration
' - startRow is the first row of a ticker that has a date value in it - rowcounter tracks the iteration number
' - Yearly_Change is the difference between the <open> value on the first iteration and the <close> value on the iteration before the ticker value changes -
' - Percent_Change is the Yearly_Change divided by the <open> value - e is the end of the current iteration

Dim total, i, startRow, rowcounter, Yearly_Change, Percent_Change, e As Double

' - Print the column headings at the top of the Summary table

Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

' - Set the initial value of the variables

e = 0
total = 0
Change = 0
Start = 2

' - Retrieve the row number of the last row with data in each worksheet by going to the bottom row and counting up until there is data

rowcounter = Cells(Rows.Count, "A").End(xlUp).Row

' - Set up the For loop

For i = 2 To rowcounter

' - Using this condition collect the values when the iteration moves to the next ticker

If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

' - Store the Total stock volume for each ticker row results in variables

total = total + Cells(i, 7).Value

' - Look after tickers with no stock volumes

    If total = 0 Then
        
' - Print the results

    Range("I" & 2 + e).Value = Cells(i, 1).Value
    Range("J" & 2 + e).Value = 0
    Range("K" & 2 + e).Value = "%" & 0
    Range("L" & 2 + e).Value = 0
    Else
            
' - Or locate the first ticker with a stock volume

    If Cells(Start, 3) = 0 Then
        
        For find_value = Start To i
            If Cells(find_value, 3).Value <> 0 Then
                     
                        Start = find_value
                        Exit For
                    End If
                 Next find_value
            End If

' - Calculate the Yearly_Change
            
            Yearly_Change = (Cells(i, 6) - Cells(Start, 3))
            
            Percent_Change = Round((Yearly_Change / Cells(Start, 3) * 100), 2)

' - Then start on the next ticker
            
            Start = i + 1

' - Print the values under the approprate column headings
            
            Range("I" & 2 + e).Value = Cells(i, 1).Value
            Range("J" & 2 + e).Value = Round(Yearly_Change, 2)
            Range("K" & 2 + e).Value = "%" & Percent_Change
            Range("L" & 2 + e).Value = total


' - Conditional formatting to fill positive values in green and negative values in red, cells with 0 value are clear

            Select Case Yearly_Change
                Case Is > 0
                Range("J" & 2 + e).Interior.ColorIndex = 4
                Case Is < 0
                    Range("J" & 2 + e).Interior.ColorIndex = 3
                Case Else
                    Range("J" & 2 + e).Interior.ColorIndex = 0
            End Select

        End If


' - Reset the variable vaules at the start of a new iteration
        
        total = 0
        Change = 0
        e = e + 1
                


' - Or while the ticker is the same keep adding up the Total Stock Volume
   
    Else

        total = total + Cells(i, 7).Value

    End If


Next i

' - Freeze top row, set titles in bold and AutoFit columns


    ActiveWindow.FreezePanes = True
    Rows("1:1").Select
    Selection.Font.Bold = True
    Cells.EntireColumn.AutoFit

End Sub

