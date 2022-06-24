Attribute VB_Name = "Module1"
Sub Main()

    ' variables
    ' ------------------------------------------------------------
    Dim sheet_count As Integer
    Dim sheet_header_row As Integer
    Dim tick_row_start As Integer
    Dim tick_row_count As Double
    Dim last_tick As String
    Dim color_idx_green As Integer
    Dim color_idx_red As Integer

    Dim curr_tick_col As Integer
    Dim curr_open_col As Integer
    Dim curr_close_col As Integer
    Dim curr_vol_col As Integer
    Dim curr_tick As String
    Dim curr_open As Double
    Dim curr_close As Double
    Dim curr_vol As Double

    Dim summ_tick_col As Integer
    Dim summ_year_delt_col As Integer
    Dim summ_percent_delt_col As Integer
    Dim summ_vol_col As Integer
    Dim summ_row As Integer
    Dim summ_row_start As Integer
    Dim summ_tick As String
    Dim summ_year_delt As Double
    Dim summ_percent_delt As Double
    Dim summ_vol As Double
    Dim summ_open As Double

    Dim great_header_col As Integer
    Dim great_tick_col As Integer
    Dim great_val_col As Integer
    Dim great_increase_row As Integer
    Dim great_decrease_row As Integer
    Dim great_vol_row As Integer
    Dim great_increase_tick As String
    Dim great_increase_val As Double
    Dim great_decrease_tick As String
    Dim great_decrease_val As Double
    Dim great_vol_tick As String
    Dim great_vol_val As Double
      
    ' constants
    ' ------------------------------------------------------------
    sheet_header_row = 1
    tick_row_start = 2
    color_idx_green = 4
    color_idx_red = 3

    curr_tick_col = 1
    curr_open_col = 3
    curr_close_col = 6
    curr_vol_col = 7

    summ_row_start = 2
    summ_tick_col = 9
    summ_year_delt_col = 10
    summ_percent_delt_col = 11
    summ_vol_col = 12

    great_header_col = 15
    great_tick_col = 16
    great_val_col = 17
    great_increase_row = 2
    great_decrease_row = 3
    great_vol_row = 4
    
    ' initializations
    ' ------------------------------------------------------------
    sheet_count = Sheets.Count()
    
    ' loop through worksheets
    For wksht = 1 To sheet_count
        
        ' worksheet headers
        Sheets(wksht).Cells(sheet_header_row, summ_tick_col).Value = "Ticker"
        Sheets(wksht).Cells(sheet_header_row, summ_year_delt_col).Value = "Yearly Change"
        Sheets(wksht).Cells(sheet_header_row, summ_percent_delt_col).Value = "Percent Change"
        Sheets(wksht).Cells(sheet_header_row, summ_vol_col).Value = "Total Stock Volume"
        
        Sheets(wksht).Cells(sheet_header_row, great_tick_col).Value = "Ticker"
        Sheets(wksht).Cells(sheet_header_row, great_val_col).Value = "Value"
        Sheets(wksht).Cells(great_increase_row, great_header_col).Value = "Greatest % Increase"
        Sheets(wksht).Cells(great_decrease_row, great_header_col).Value = "Greatest % Decrease"
        Sheets(wksht).Cells(great_vol_row, great_header_col).Value = "Greatest Total Volume"
       
        ' worksheet initialization
        summ_row = summ_row_start
        great_increase_val = 0
        great_decrease_val = 0
        great_vol_val = 0
        
        ' loop through ticks
        tick_row_count = Sheets(wksht).Cells(Rows.Count, curr_tick_col).End(xlUp).Row
        
        For tick = tick_row_start To tick_row_count
        
            ' grab current values
            curr_tick = Sheets(wksht).Cells(tick, curr_tick_col).Value
            curr_open = Sheets(wksht).Cells(tick, curr_open_col).Value
            curr_close = Sheets(wksht).Cells(tick, curr_close_col).Value
            curr_vol = Sheets(wksht).Cells(tick, curr_vol_col).Value
            next_tick = Sheets(wksht).Cells(tick + 1, curr_tick_col).Value

            ' initialize first tick
            If tick = tick_row_start Then
                summ_open = curr_open
                summ_tick = curr_tick
                summ_vol = 0
                last_tick = curr_tick
                Sheets(wksht).Cells(summ_row, summ_tick_col).Value = curr_tick
            End If

            ' tabulate summaries
            If curr_tick <> last_tick Then
                summ_row = summ_row + 1
                summ_open = curr_open
                summ_tick = curr_tick
                summ_vol = curr_vol
                last_tick = curr_tick
                Sheets(wksht).Cells(summ_row, summ_tick_col).Value = curr_tick
            Else
                summ_vol = summ_vol + curr_vol
                summ_year_delt = curr_close - summ_open
                summ_percent_delt = summ_year_delt / summ_open
            End If
                
            ' update worksheet and identify greatest
            If next_tick <> curr_tick Then
            
                Sheets(wksht).Cells(summ_row, summ_year_delt_col).Value = summ_year_delt
                Sheets(wksht).Cells(summ_row, summ_percent_delt_col).Value = summ_percent_delt
                Sheets(wksht).Cells(summ_row, summ_vol_col).Value = summ_vol
                    
                If summ_percent_delt > great_increase_val Then
                    great_increase_tick = summ_tick
                    great_increase_val = summ_percent_delt
                    Sheets(wksht).Cells(great_increase_row, great_tick_col).Value = great_increase_tick
                    Sheets(wksht).Cells(great_increase_row, great_val_col).Value = great_increase_val
                End If

                If summ_percent_delt < great_decrease_val Then
                    great_decrease_tick = summ_tick
                    great_decrease_val = summ_percent_delt
                    Sheets(wksht).Cells(great_decrease_row, great_tick_col).Value = great_decrease_tick
                    Sheets(wksht).Cells(great_decrease_row, great_val_col).Value = great_decrease_val
                End If

                If summ_vol > great_vol_val Then
                    great_vol_tick = summ_tick
                    great_vol_val = summ_vol
                    Sheets(wksht).Cells(great_vol_row, great_tick_col).Value = great_vol_tick
                    Sheets(wksht).Cells(great_vol_row, great_val_col).Value = great_vol_val
                End If
                
                ' highlight positive and negative yearly change
                If summ_year_delt < 0 Then
                    Sheets(wksht).Cells(summ_row, summ_year_delt_col).Interior.ColorIndex = color_idx_red
                ElseIf summ_year_delt > 0 Then
                    Sheets(wksht).Cells(summ_row, summ_year_delt_col).Interior.ColorIndex = color_idx_green
                End If
                
            End If
            
        Next tick
    
    Next wksht

End Sub

