Attribute VB_Name = "Module1"
Sub all_procedures()
    
    Call print_ticker_name_list
    Call yearly_change
    Call yearly_change_percentage
    Call yearly_volume_usage
      
End Sub


Sub print_ticker_name_list()

    Dim ws As Worksheet
    Dim ticker As String
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim total_stock_volume As Integer
    Dim last_row As Long
    Dim summary_ticker As Integer
    Dim summary_change As Integer
    
    For Each ws In Worksheets
    ws.Cells(1, 9).Value = "Ticker"
    
            'defined last row and where to start the list of ticker
            summary_ticker = 2
            last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
              
            'loop through ticker list
            For i = 2 To last_row
            
                'defining where to stop between each ticker
                If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                ticker = ws.Cells(i, 1).Value
                
                'print ticker list
                ws.Range("I" & summary_ticker).Value = ticker
                
                'giving interval to print each ticker
                summary_ticker = summary_ticker + 1
                        
                      
                End If
                
                Next i
                
        Next ws
                
End Sub



Sub yearly_change()

    Dim ws As Worksheet
    Dim ticker As String
    Dim yearly_change As Double
    Dim yearly_opening As Double
    Dim percent_change As Double
    Dim total_stock_volume As Integer
    Dim last_row As Long
    Dim summary_ticker As Integer
    Dim summary_change As Integer
    Dim total_volume As Double
    Dim first_row_open As Double
    Dim last_row_close As Double

    For Each ws In Worksheets
    ws.Cells(1, 10).Value = "Yearly Change"
    summary_change = 2
    last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    total_volume = 0
    first_row_open = ws.Cells(2, 3).Value
    
                    'loop through total volume
                For i = 2 To last_row
                    If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                            last_row_close = ws.Cells(i, 6).Value
                            
                            yearly_change = last_row_close - first_row_open
                            
                            first_row_open = ws.Cells(i + 1, 3).Value
                    
                    ws.Range("J" & summary_change).Value = yearly_change
                        If yearly_change > 0 Then
                            ws.Range("J" & summary_change).Interior.ColorIndex = 4
                                Else
                                ws.Range("J" & summary_change).Interior.ColorIndex = 3
                    
                            End If
                    
                    summary_change = summary_change + 1
                    
                    'total_volume = 0
                    
                    Else
                    
                   'total_volume = total_volume + Cells(i, 7).Value
                   
                    End If
                    
                Next i
                
        Next ws
           


End Sub

Sub yearly_change_percentage()

    Dim ws As Worksheet
    Dim first_open As Long
    Dim last_close As Long
    Dim last_row As Long
    Dim year_change_percentage As Long
    Dim summary_change As Double
    
    For Each ws In Worksheets
    ws.Cells(1, 11).Value = "Yearly Change %"
    first_open = ws.Cells(2, 3).Value
    last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    summary_change = 2
    
                For i = 2 To last_row
                    If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                        last_close = ws.Cells(i, 6).Value
                            If first_open = 0 Then
                                ws.Range("K" & summary_change).Value = "Null"
                            Else
                                
                            year_change_percentage = CLng((((last_close - first_open) / first_open) * 100))
                                                
                           End If
                            
                        first_open = ws.Cells(i + 1, 3).Value
                        
                        ws.Range("K" & summary_change).Value = year_change_percentage
                        
                        summary_change = summary_change + 1
                        
                        Else
                        
                        End If
                        
                    Next i
        
        Next ws
        
End Sub


Sub yearly_volume_usage()

    Dim ws As Worksheet
    Dim ticker As String
    Dim yearly_change As Double
    Dim yearly_opening As Double
    Dim yearly_close As Double
    Dim percent_change As Double
    Dim total_stock_volume As Integer
    Dim last_row As Long
    Dim summary_ticker As Integer
    Dim summary_change As Integer
    Dim total_volume As Double

    For Each ws In Worksheets
    ws.Cells(1, 12).Value = "Total Volume"
    summary_change = 2
    last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    total_volume = 0
        'loop through total volume
                    For i = 2 To last_row
                        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                                
                        
                        ws.Range("L" & summary_change).Value = total_volume
                        
                        
                        summary_change = summary_change + 1
                        
                        total_volume = 0
                        
                        Else
                        
                       total_volume = total_volume + ws.Cells(i, 7).Value
                       
                        End If
                        
                    Next i

        Next ws
        
End Sub

