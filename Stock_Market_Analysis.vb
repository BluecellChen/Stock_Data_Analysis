Sub Stock_Market_Analysis():

    ' Declare variables to hold values
    Dim Ticker As String
    Dim Total_Stock_Volume As Double
    Dim lastrow As Long
    Dim open_price As Double
    Dim close_price As Double
    Dim Yearly_Change As Double
    Dim count As Long
    Dim yearopenrow As Long
    Dim Percent_Change As Double
    Dim Greatest_Increase As Double
    Dim Greatest_Decrease As Double
    Dim Greatest_Total_Volume As Double
    
    

    ' LOOP THROUGH ALL SHEETS
    For Each ws In Worksheets
        ' Headers   
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"

        ' Set initial values of variables. 
        Total_Stock_Volume = 0
        count = 2
        yearopenrow = 2
        open_price = ws.Cells(2, 3).Value
        Yearly_Change = 0
        Percent_Change = 0
        'Greatest_Total_Volume = 0
        'Greatest_Increase =0'
        'Greatest_Decrease = 0

        ' Determine the last row
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' In each sheet, loop through from row 2 to lastrow
        For i = 2 to lastrow
            ' easy ----------------------------------------------------
            ' Check if Ticker is different in the next row
            If ws.Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                'set current row ticker as Ticker 
                Ticker = ws.Cells(i, 1).Value
                ' Put the Ticker and Total_Stock_Volume values in the Table
                ws.Range("I" & count).Value = Ticker
                ws.Range("L" & count).Value = Total_Stock_Volume
                ' Reset Total_Stock_Volume
                Total_Stock_Volume = 0

            'moderate---------------------------------------------------

                'set yearly open, yearly close 
                open_price = ws.Range("C" & yearopenrow)
                close_price = ws.Range("F" & i)
                Yearly_Change = close_price - open_price
                ' put Yearly_Change value in the Table
                ws.Range("J" & count).Value = Yearly_Change

                ' Determine Percent Change
                If open_price = 0 and Yearly_Change = 0 Then 
                    Percent_Change = 0
                ElseIf open_price = 0 and close_price <> 0 Then 
                    Percent_Change = 1
                Else 
                    Yearly_Change = close_price - open_price
                    open_price = ws.Range("C" & yearopenrow)
                    Percent_Change = Yearly_Change / Open_price
                End If

                ' % Formatting of percentage
                ws.Range("K" & count).NumberFormat = "0.00%"
                ws.Range("K" & count).Value = Percent_Change

                ' Conditional formatting of Yearly_Change column
                If ws.Range("J" & count).Value >= 0 Then
                    ws.Range("J" & count).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & count).Interior.ColorIndex = 3
                End If
            
                ' reset 
                Open_Price = ws.Cells(i + 1, 3)    
                count = count + 1
                yearopenrow = i + 1

            Else  ' If cells have the same ticker, Add current row Volume to Total_Stock_Volume
            Total_Stock_Volume = Total_Stock_Volume + ws.Cells(I, 7).Value
            End If
        Next i
        
        'hard ---------------------------------------------------------------------------
        ' Greatest % Increase, Greatest % Decrease and Greatest Total Volume
            lastrow = ws.Cells(Rows.Count, 11).End(xlUp).Row
        For i = 2 To lastrow
            If ws.Range("K" & i).Value > ws.Range("Q2").Value Then
                ws.Range("Q2").Value = ws.Range("K" & i).Value
                ws.Range("P2").Value = ws.Range("I" & i).Value
            End If

            If ws.Range("K" & i).Value < ws.Range("Q3").Value Then
                ws.Range("Q3").Value = ws.Range("K" & i).Value
                ws.Range("P3").Value = ws.Range("I" & i).Value
            End If

            If ws.Range("L" & i).Value > ws.Range("Q4").Value Then
                ws.Range("Q4").Value = ws.Range("L" & i).Value
                ws.Range("P4").Value = ws.Range("I" & i).Value
            End If

            ws.Range("Q2").NumberFormat = "0.00%"
            ws.Range("Q3").NumberFormat = "0.00%"
        Next i
        
    Next ws

End Sub