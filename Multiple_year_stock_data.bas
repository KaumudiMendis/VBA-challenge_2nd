Attribute VB_Name = "Module1"
Sub VBAProject_Alphabetical_testing()

' --------------------------------------------
' LOOP THROUGH ALL SHEETS
' --------------------------------------------

'Create variable of the worksheet to run the code in all work sheet at once in the workbook

Dim ws As Worksheet
For Each ws In Worksheets

'----------------------------------
'Create all  variables for Calculation and output: Ticker_Symbol, Opening_Price,Closing_Price,Percent_change,Total_Stock_Volume
'----------------------------------

Dim Ticker_Symbol As String
Dim Opening_Price As Double
Dim Closing_Price As Double
Dim Yearly_Change As Double
Dim Total_Stock_Volume As Double
Dim Percent_Change As Double

'Create a variable to set up a row to start

Dim start_data As Integer

    'Set column headers for each worksheet

    ws.Cells(1, 9).Value = "Ticker Symbol"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"

    'Set intigers for the loop to start
    start_data = 2
    previous_i = 1
    Total_Stock_Volume = 0

    'Go to the last row of coumn "A" for each worksheet(ws)

    EndRow = ws.Cells(Rows.Count, "A").End(xlUp).Row

' --------------------------------------------
        'CALCULATING YEARLY CHANGE,PERCENT CHANGE AND TOTAL STOCK VOLUME FOR EACH TICKER SYMBOL
' --------------------------------------------

    For i = 2 To EndRow

        'Until the Ticker Symbol change run below and Get the Ticker Symbol
            
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            Ticker_Symbol = ws.Cells(i, 1).Value

            'Intiate the variable to go to the next Ticker Symbol Alphabet

            previous_i = previous_i + 1

            '-------
            'For each ticker symbol,
            '1. Opening Price: on the first date/row start get the value from Column 3 ("C")
            '2. Closing Price:  for the last day/last row get the value from column 6 ("F")
            '3. Total stock volume: total for each ticker symbol from column 7("G") and calculated using a separate for loop.
            

            Opening_Price = ws.Cells(previous_i, 3).Value
            Closing_Price = ws.Cells(i, 6).Value

                For j = previous_i To i
                    Total_Stock_Volume = Total_Stock_Volume + ws.Cells(j, 7).Value
                Next j
            
               
            
                'When the Opening Price is zero for a particular Ticker Symbol

                If Opening_Price = 0 Then
                    Percent_Change = Closing_Price

                Else
                    Yearly_Change = Closing_Price - Opening_Price

                    Percent_Change = Yearly_Change / Opening_Price

                End If
         '--------------------------------------------------

            'Results for summary table 1:  Ticker Symbol, Yearly Change,Percent Change, Total Stock Volume

            ws.Cells(start_data, 9).Value = Ticker_Symbol
            ws.Cells(start_data, 10).Value = Yearly_Change
            ws.Cells(start_data, 11).Value = Percent_Change
            ws.Cells(start_data, 11).NumberFormat = "0.00%"
            ws.Cells(start_data, 12).Value = Total_Stock_Volume

            'Get next row data for summary table 1.

            start_data = start_data + 1
            Total_Stock_Volume = 0
            Yearly_Change = 0
            Percent_Change = 0

        'start the loop again.
        
        previous_i = i

        End If
        
    Next i
'------------------------------
'CONDITIONAL FORMATING
'------------------------------
'The end row for column J

    jEndRow = ws.Cells(Rows.Count, "J").End(xlUp).Row


        For j = 2 To jEndRow

            'if greater than or less than zero
            If ws.Cells(j, 10) > 0 Then

                ws.Cells(j, 10).Interior.ColorIndex = 4

            Else

                ws.Cells(j, 10).Interior.ColorIndex = 3
            End If

        Next j
'---------------------------------------------
      'TABLE 2
'---------------------------------------------


'-------------------------------------------------------------------
'CALCULATING   "Greatest % increase", "Greatest % decrease" and "Greatest total volume" IN TABLE 2
'-------------------------------------------------------------------

    'Go through column K
    
    kEndRow = ws.Cells(Rows.Count, "K").End(xlUp).Row

    'Create variable for Summary Table 2 values

    Increase = 0
    Decrease = 0
    Greatest = 0

        'Find the maximum and minimum for percentage change and the maximum volume
        
        For k = 3 To kEndRow

            'Define current row for percentage and previous row for percentage increment to check
            
            last_k = k - 1
            current_k = ws.Cells(k, 11).Value
            previous_k = ws.Cells(last_k, 11).Value

            'greatest total volume row and previous greatest volume row
            volume = ws.Cells(k, 12).Value
            previous_vol = ws.Cells(last_k, 12).Value
            
            
            'Find the Greatest % increase and   define name for increase percentage
            
            If Increase > current_k And Increase > previous_k Then
                Increase = Increase

            ElseIf current_k > Increase And current_k > previous_k Then

                Increase = current_k
                increase_name = ws.Cells(k, 9).Value

            ElseIf previous_k > Increase And previous_k > current_k Then

                Increase = previous_k
                increase_name = ws.Cells(last_k, 9).Value

            End If


       '--------------------------------------------------
            'Find the greatest % decrease and define name for decrease percentage

            If Decrease < current_k And Decrease < previous_k Then
                Decrease = Decrease

            ElseIf current_k < Increase And current_k < previous_k Then

                Decrease = current_k
                decrease_name = ws.Cells(k, 9).Value

            ElseIf previous_k < Increase And previous_k < current_k Then

                Decrease = previous_k

                decrease_name = ws.Cells(last_k, 9).Value

            End If

       '--------------------------------------------------
           'Find the greatest volume and define name for greatest volume

            If Greatest > volume And Greatest > previous_vol Then

                Greatest = Greatest

            ElseIf volume > Greatest And volume > previous_vol Then

                Greatest = volume
                greatest_name = ws.Cells(k, 9).Value

            ElseIf previous_vol > Greatest And previous_vol > volume Then

                Greatest = previous_vol
                greatest_name = ws.Cells(last_k, 9).Value

            End If

        Next k
  '--------------------------------------------------
    
' Assign names for greatest increase,greatest decrease, and  greatest volume

    ws.Range("N1").Value = "Column Name"
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest Total Volume"
    ws.Range("O1").Value = "Ticker Name"
    ws.Range("P1").Value = "Value"

    'Get for greatest increase, greatest increase, and  greatest volume Ticker name
    ws.Range("O2").Value = increase_name
    ws.Range("O3").Value = decrease_name
    ws.Range("O4").Value = greatest_name
    ws.Range("P2").Value = Increase
    ws.Range("P3").Value = Decrease
    ws.Range("P4").Value = Greatest

    'Greatest increase and decrease in percentage format

    ws.Range("P2").NumberFormat = "0.00%"
    ws.Range("P3").NumberFormat = "0.00%"




'run the code in next worksheet
Next ws

'--------------------------------------------------
End Sub

