VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub StockSummery()

Dim ticker As String
Dim year_open As Double
Dim year_close As Double
Dim yearly_change As Double
Dim total_stock_volume As Double
Dim percent_change As Double
Dim start_data As Long
Dim start_summary As Integer
Dim ws As Worksheet
Dim lastrow As Long
Dim klastrow As Integer
Dim increase As Double
Dim decrease As Double
Dim greatest_volume As Double
Dim increase_name As String
Dim decrease_name As String
Dim greatest_name As String


    For Each ws In Worksheets
        ' asign summary table headers
    ws.Cells(1, 10).Value = "Ticker"
    ws.Cells(1, 11).Value = "Yearly Change"
    ws.Cells(1, 12).Value = "Percent Change"
    ws.Cells(1, 13).Value = "Total Stock Volume"
        ' Check to make sure the whole sheet is in general format before starting
    ws.Columns("A:Z").NumberFormat = "General"
        ' assign variable start values
    start_data = 2
    start_summary = 2
    total_stock_volume = 0
    greatest_volume = 0
         ' find lastrow
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        '    MsgBox (lastrow)
    For i = 2 To lastrow
            ' look for start of next ticker
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ' Grab Ticker
        ticker = ws.Cells(i, 1).Value
            ' Grab year_open
        year_open = ws.Cells(start_data, 3).Value
            ' Grab year_close
        year_close = ws.Cells(i, 6).Value
            
            ' loop to grab total stock volume
        For j = start_data To i
                ' Find total_stock_volume
            total_stock_volume = total_stock_volume + ws.Cells(j, 7).Value
        
        Next j
             ' Check if total_stock_volume is greatest
        If total_stock_volume > greatest_volume Then
            greatest_volume = total_stock_volume
            greatest_name = ws.Cells(i, 1).Value
        End If
            ' If Else statment to avoid divide by 0 (sourced from ermiasgelaye)
        If year_open = 0 Then
            percent_change = year_close

        Else
                ' determines yearly_change and percent_change
            yearly_change = year_close - year_open
            
            percent_change = yearly_change / year_open
        
        End If
            ' paste data into summary table
        ws.Cells(start_summary, 10).Value = ticker
        ws.Cells(start_summary, 11).Value = yearly_change
        ws.Cells(start_summary, 12).Value = percent_change
        ws.Cells(start_summary, 13).Value = total_stock_volume
            ' Format percent_change as percent
        ws.Cells(start_summary, 12).NumberFormat = "0.00%"
            ' increment start_summary
        start_summary = start_summary + 1
            ' reset variables
        total_stock_volume = 0
        yearly_change = 0
        percent_change = 0
            ' move start data to start of new ticker
        start_data = i + 1
        
        End If

    Next i

klastrow = ws.Cells(Rows.Count, 11).End(xlUp).Row
increase = 0
decrease = 0

    For k = 2 To klastrow
        
        current_k = ws.Cells(k, 12).Value
        
        volume = ws.Cells(k, 13).Value
        
        previous_vol = ws.Cells(2, 12).Value
        
'--------------------------------------------------

        If increase > current_k Then
        
            increase = increase
            
        ElseIf current_k > increase Then
        
            increase = current_k
            
            increase_name = ws.Cells(k, 10).Value
            
        End If
        
'--------------------------------------------------

        If decrease < current_k Then
        
            decrease = decrease
            
        ElseIf current_k < decrease Then
        
            decrease = current_k
            
            decrease_name = ws.Cells(k, 10).Value
            
        End If
            
'--------------------------------------------------

            'Conditional Formatting Yearly Change Column
        If ws.Cells(k, 11).Value > 0 Then
        
            ws.Cells(k, 11).Interior.ColorIndex = 4
            
        Else
        
            ws.Cells(k, 11).Interior.ColorIndex = 3
        End If
             
'--------------------------------------------------
             
      Next k
        ' assign headers for summary chart 2
    ws.Cells(2, 15).Value = "Greatest Increase"
    ws.Cells(3, 15).Value = "Greatest Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
        ' write data to table
    ws.Cells(2, 16).Value = increase_name
    ws.Cells(2, 17).Value = increase
    ws.Cells(3, 16).Value = decrease_name
    ws.Cells(3, 17).Value = decrease
    ws.Cells(4, 16).Value = greatest_name
   ws.Cells(4, 17).Value = greatest_volume
        
    ws.Range("Q2:Q3").NumberFormat = "0.00%"
        
'--------------------------------------------------
        
        ' Autofit whole sheet
    ws.Columns("A:Z").AutoFit
    
    Next ws
    
End Sub

'made with reference from https://github.com/ermiasgelaye/VBA-challenge
'nothing was full sale copy pasted
'other than the one line that has a reference in the note their code was mostly used to debug my own


