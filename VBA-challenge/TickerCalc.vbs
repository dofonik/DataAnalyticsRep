Attribute VB_Name = "Module1"
Sub TickerCalc():

    'Dom Foale - PSUEDOCODE/PLAN/NOTES
    'FOR all sheets in the excel file
    'FOR each ticker -> need counter to track starting row for volume
    'RECORD ticker symbol
    'DETERMINE last row for FOR loop
    'DETERMINE where last row is for this ticker -> condition to check if next row ticker is different
    'RECORD first open -> use counter for starting row
    'RECORD last close
    'RECORD sum of all volume -> using long variable still not large enough for some instances (counter required)
    'CALCULATE yearly change, percentage change
    'OUTPUT ticker, yearly change, percentage change, sum of volume
    'SET format and colour for cells
    'END FOR each ticker
    'FOR all rows of output table
    'Record highest/lowest/greatest variables using conditional
    'END FOR each output table row
    'END FOR each sheet
    
    'Avoid crashing
    Application.ScreenUpdating = False
    
    'For loop to loop through all sheets in workbook
    For Each ws In Worksheets

        'Define all required variables
        Dim WorksheetName As String
        Dim LastRowInt As Long 'Integer not large enough
        Dim i As Long 'Integer not large enough
        Dim YearOpen As Double
        Dim YearClose As Double
        Dim YearChange As Double
        Dim PercentChange As Double
        Dim OutputTableRow As Integer
        Dim TickStartRow As Long 'Integer not large enough
        
        'Set text titles for this worksheet
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        'Determine last row of this worksheet
        LastRowInt = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Set OutputTableRow and TickStartRow to 2
        OutputTableRow = 2
        TickStartRow = 2
        
        For i = 2 To LastRowInt
            
            'Only trigger once new ticker is found in the next row of current i row
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                
                'Record ticker data
                YearOpen = ws.Cells(TickStartRow, 3).Value
                YearClose = ws.Cells(i, 6).Value
                                
                'Calculate outputs for output table
                YearChange = YearClose - YearOpen
                PercentChange = ((YearClose - YearOpen) / YearOpen)
                
                'Output all required data to output table
                ws.Cells(OutputTableRow, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(OutputTableRow, 10).Value = YearChange
                ws.Cells(OutputTableRow, 11).Value = Format(PercentChange, "Percent") 'Format cell while outputting
                'Ouput volume with sum function to avoid long overflow
                ws.Cells(OutputTableRow, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(TickStartRow, 7), ws.Cells(i, 7)))
                
                'Conditionally format percentage change cell
                If ws.Cells(OutputTableRow, 10) < 0 Then
                    
                    'Set colour to red if negative
                    ws.Cells(OutputTableRow, 10).Interior.ColorIndex = 3
                    
                Else
                
                    'Else value is positive and set to green
                    ws.Cells(OutputTableRow, 10).Interior.ColorIndex = 4
                    
                End If
                
                'Increment output table row and update ticker start row for volume
                OutputTableRow = OutputTableRow + 1
                TickStartRow = i + 1
                
                
            End If
            
        Next i
        
        'Set output cells to 0 in case there is a value already there
        ws.Range("Q2").Value = 0
        ws.Range("Q3").Value = 0
        ws.Range("Q4").Value = 0
        
        'BONUS For loop to determine greatest percentage increase, greatest percentage decrease and greatest volume
        For i = 2 To (OutputTableRow - 1)
            
            'Check if current greatest values are less than looping row, if so replaces greatest value
            'Greatest % Increase
            If ws.Cells(i, 11).Value > ws.Range("Q2").Value Then
                
                ws.Range("Q2").Value = Format(ws.Cells(i, 11).Value, "Percent")
                'Grab ticker
                ws.Range("P2").Value = ws.Cells(i, 9).Value
                
            End If
            'Greatest % Decrease
            If ws.Cells(i, 11).Value < ws.Range("Q3").Value Then
                
                ws.Range("Q3").Value = Format(ws.Cells(i, 11).Value, "Percent")
                ws.Range("P3").Value = ws.Cells(i, 9).Value
                
            End If
            'Greatest Total Volume
            If ws.Cells(i, 12).Value > ws.Range("Q4").Value Then
                
                ws.Range("Q4").Value = ws.Cells(i, 12).Value
                ws.Range("P4").Value = ws.Cells(i, 9).Value
                
            End If
            
        Next i
    
    'Autofit columns for ease of viewing
    WorksheetName = ws.Name
    Worksheets(WorksheetName).Columns("A:Z").AutoFit
    
    Next ws
    
    Application.ScreenUpdating = True
    
End Sub
