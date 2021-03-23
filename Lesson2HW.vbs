Sub TickerProject3()

    ' ----------------  Define my variables
    ' Define the variable that will contain the ticker symbol in turn
    Dim str_ticker_symbol As String
    ' Define the variable that will contain the opening price and the closing price of each ticker symbol
    Dim dbl_opening_price As Double
    Dim dbl_closing_price As Double
    ' Variable that will contain the yearly change and percent change between closing and opening prices
    Dim dbl_yearly_change As Double
    Dim dbl_percent_change As Double
    Dim eval_change as integer
    ' Variable that will add up the total stock volume
    Dim dbl_total_volume As Double
    dbl_total_volume = 0 ' Not necessary
    ' Worksheet variables
    Dim I as Long  ' First loop variable - complete list
    Dim J as Long  ' Second loop variable - summary table
    Dim ws As Worksheet 
    WS_Count = ActiveWorkbook.Worksheets.Count
    'Row Variables
    Dim row_complete_table As Long ' Apparently unnecesary
    Dim row_summary As Long
    Dim last_row As Long
    Dim last_row_sum as long
    ' Variables that will hold greater % increase % decrease and greater total volume
    Dim dbl_p_increase as Double
    dbl_p_increase = 0
    Dim dbl_p_decrease as Double
    dbl_p_decrease = 0
    Dim dbl_greater_tot_vol as Double
    dlb_greater_tot_vol = 0
    dim str_ticker_inc as String
    dim str_ticker_dec as String
    dim str_ticker_hi_vol as String

    ' First we need to create titles in every sheet and apply format

    For Each ws In ThisWorkbook.Worksheets

        row_summary = 2

        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ' Bonus
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("o2").Value = "Greatest % Increase"
        ws.Range("o3").Value = "Greatest % Decrease" 
        ws.Range("o4").Value = "Greatest Total Volume"  
        ws.Columns("O:O").AutoFit
        ' Formats    
        ws.Range("K:K").NumberFormat = "0.00%"
        ws.Range("I1:L1").HorizontalAlignment = xlCenter
        ws.Range("I1:L1").VerticalAlignment = xlCenter
        ws.Range("I1:L1").Font.FontStyle = "Bold Italic"
        ws.Range("I1:L1").Font.Color = vbWhite
        ws.Range("I1:L1").Interior.Color = vbBlack
        ws.Range("I1:L1").WrapText = True

        ' We need create a loop that identifies the first row with a new ticker symbol and save the opening price
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row

        For I = 2 To last_row

            ' Is it the first row? Set up the condition to identify first row of each ticker
            If ws.Cells(I - 1, 1).Value <> ws.Cells(I, 1).Value Then

                ' Obtain ticker symbol and opening price  of the first row
                str_ticker_symbol = ws.Cells(I, 1).Value
                dbl_opening_price = ws.Cells(I, 3).Value

                ' Reset volume and start adding the volume in each line
                dbl_total_volume = ws.Cells(I, 7).Value

            ' Is it the last row? Set up the condition to identify the last row of each ticker
            elseif ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then

                ' Obtain the closing price
                dbl_closing_price = ws.Cells(I, 6)

                ' *************Time to paste all values
                ' First we paste the ticker symbol
                ws.Cells(row_summary, 9).Value = str_ticker_symbol

                ' Then we calculate and paste yearly change
                dbl_yearly_change = dbl_closing_price - dbl_opening_price

                ws.Cells(row_summary, 10).Value = dbl_yearly_change

                ' Next we calculate and paste the percent change
                dbl_percent_change = ((dbl_closing_price - dbl_opening_price) - 1)
                ws.Cells(row_summary, 11).Value = dbl_percent_change

                    if dbl_percent_change > 0 Then
                        ws.Cells(row_summary, 11).Interior.Color = vbGreen
                    elseif dbl_percent_change < 0 Then
                        ws.Cells(row_summary, 11).Interior.Color = vbRed
                    end if

                ' Finally we paste total stock volume
                ws.Cells(row_summary, 12).Value = dbl_total_volume

                ' Add volume
                dbl_total_volume = dbl_total_volume + ws.Cells(I, 7).Value

                ' Add one line to row summary
                row_summary = row_summary + 1

            Else
                dbl_total_volume = dbl_total_volume + ws.Cells(I, 7).Value

            End If
        Next I

        last_row_sum = ws.Cells(Rows.Count, 9).End(xlUp).Row

        dbl_p_increase = 0
        dbl_p_decrease = 0
        dbl_total_volume = 0

        For J = 2 to last_row_sum

            If ws.cells(j,11) > dbl_p_increase Then
            dbl_p_increase = ws.cells(j,11).value
            str_ticker_inc = ws.cells(j,9).value
            ws.range("p2").value = str_ticker_inc
            ws.range("q2").value = dbl_p_increase
            ws.Range("q2").NumberFormat = "0.00%"

            End If

            If ws.cells(j,11) < dbl_p_decrease Then
            dbl_p_decrease = ws.cells(j,11).value
            str_ticker_dec = ws.cells(j,9).value
            ws.range("p3").value = str_ticker_dec
            ws.range("q3").value = dbl_p_decrease
            ws.Range("q3").NumberFormat = "0.00%"

            End If

            If ws.cells(j,12) > dbl_total_volume Then
            dbl_total_volume = ws.cells(j,12).value
            str_ticker_hi_vol = ws.cells(j,9).value
            ws.range("p4").value = str_ticker_hi_vol
            ws.range("q4").value = dbl_total_volume
            ws.Range("q4").NumberFormat = Number

            End if

        Next J

    Next ws

End Sub
