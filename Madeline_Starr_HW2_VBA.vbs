'Part I: Create of list of all unique ticker names in each worksheet
'-----------------------------------------------------------

'assign initial variables
'-----------------------------------------------------------
Sub stockticker()
Dim stock As Integer
Dim LastRow As Long
Dim stocksave As String
Dim ws As Worksheet
'Set first_sheet = Worksheets("A")


'cause all subsequent code to initialize on all pages in a workbook
'-----------------------------------------------------------
For Each ws In Worksheets
    
    'assign values to column and row headers
    '-----------------------------------------------------------
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
    'set stock value to 2 for second row
    '-----------------------------------------------------------
    stock = 2
        
    'find value of the last row, assign it to a variable and loop from 2 to LastRow
    '-----------------------------------------------------------
    LastRow = ws.Range("A" & Rows.Count).End(xlUp).Row
    
    For i = 2 To LastRow
        
        'Once code hits a new <ticker> value, copy that value into the Ticker column
        'otherwise, ignore the <ticker> value
        '-----------------------------------------------------------
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            stocksave = ws.Cells(i, 1).Value
            ws.Range("i" & stock).Value = stocksave
            
            'increase counter to move down a row after new entry
            '-----------------------------------------------------------
            stock = stock + 1
        End If
    Next i

'Part II:
    'a: Create a volume total for each unique ticker (easy)
    'b: Find the difference between opening price on the first day and closing price on the last day (moderate)
    'c: Create conditional formating to indicate positive or negative changes (moderate)
    'd: Calculate % change opening price on the first day and closing price on the last day (moderate)
        '1: Watch out for divide by zeros!
'-----------------------------------------------------------

    'Start a new set of variables
    '-----------------------------------------------------------
    Dim stockmove As Integer
    Dim runtotal As Double
    Dim startopen As Double
    Dim endclose As Double
    Dim percentchange As Double
    Dim runtotalmax As Double
    Dim percentchangemin As Double
    Dim percentchangemax As Double
    Dim tickerrtm As String
    Dim tickerpcmax As String
    Dim tickerpcmin As String
    
    runtotalmax = 0
    percentchangemin = 0
    percentchangemax = 0
    
    'stockmove represents the row of the cell that will be populated
    '-----------------------------------------------------------
    stockmove = 2
    
    'Because of the logic used, need to end at LastRow + 1 to avoid incomplete data population
    '-----------------------------------------------------------
    For j = 2 To LastRow + 1
    
        'The first row is a special case, because there is nothing above it
        '-----------------------------------------------------------
        If j = 2 Then
            startopen = ws.Range("C" & j)
        End If
        
        'Part II a: Volume calculation
        'if <ticker> column value is the same as the Ticker value, then start a running total of the <vol> column
        '-----------------------------------------------------------
        If ws.Cells(j, 1).Value = ws.Range("I" & stockmove).Value Then
            runtotal = ws.Range("G" & j).Value + ws.Range("L" & stockmove).Value
            ws.Range("L" & stockmove).Value = runtotal
            
        'Part II b, c, and d: calculate Yearly Change, impliment conditional formating,
        'and calculate percent change
        '-----------------------------------------------------------
        ElseIf ws.Cells(j, 1).Value <> ws.Range("I" & stockmove).Value Then
            
            endclose = ws.Range("F" & j - 1).Value
            ws.Range("J" & stockmove).Value = endclose - startopen
                If ws.Range("J" & stockmove).Value >= 0 Then
                    ws.Range("J" & stockmove).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & stockmove).Interior.ColorIndex = 3
                End If
 
            'calculate the percent change
            'This treats the case were the starting open price is 0 (circumvents dividing by 0)
            '-----------------------------------------------------------
            If startopen = 0 Then
                percentchange = "0"
                
            'For all other cases
            '-----------------------------------------------------------
            Else
                percentchange = ws.Range("J" & stockmove).Value / startopen
                ws.Range("K" & stockmove).Value = Format(percentchange, "0.00%")
                
                'if <ticker> column value is the same as the Ticker value, add one to stockmove (row count for Ticker list)
                'and save just the volume of that row to G&i before looping
                stockmove = stockmove + 1
                runtotal = ws.Range("G" & j).Value
            End If
                If ws.Range("I" & stockmove) <> "" Then
                    ws.Range("L" & stockmove).Value = runtotal
                End If
            startopen = ws.Range("C" & j)
  
        End If
        
'Part III: Find the Max total volume, Max percent change, and Min percent change
'and copy those values, along with the ticker name into new cells (Hard)
'-----------------------------------------------------------
        
        If runtotal > runtotalmax Then
            runtotalmax = runtotal
            tickerrtm = ws.Range("I" & stockmove)
        End If
        
        If percentchange > percentchangemax Then
            percentchangemax = percentchange
            tickerpcmax = ws.Range("I" & stockmove - 1)
        End If
        
        If percentchange < percentchangemin Then
            percentchangemin = percentchange
            tickerpcmin = ws.Range("I" & stockmove - 1)
        End If
        
    Next j
ws.Range("P2").Value = tickerpcmax
ws.Range("Q2").Value = Format(percentchangemax, "0.00%")
ws.Range("P3").Value = tickerpcmin
ws.Range("Q3").Value = Format(percentchangemin, "0.00%")
ws.Range("P4").Value = tickerrtm
ws.Range("Q4").Value = runtotalmax


Next ws

'Message that the code has finished running for a quick visual check
'-----------------------------------------------------------
MsgBox ("All Done!!")
End Sub



'This code is to copy all of the created data from all sheets onto the first sheet
'insert after "Next j" and uncomment the variable Set first_sheet = Worksheets("A")
'to make this function
'-----------------------------------------------------------

'LastRowI = ws.Range("I" & Rows.Count).End(xlUp).Row
    'For j = 2 To LastRowI
    'If ws.Cells(j, 1).Value = ws.Cells(j + 1, 1).Value Then
        
    
    'For every sheet that isn't "A", find the last row of column I on the current sheet
    'and the last row of column I on sheet "A"
    'Then copy the information from columns I and J in the current sheets to the bottom of
    'I and J in sheet "A"
    'If ws.Name <> "A" Then
        'LastRowI = ws.Range("I" & Rows.Count).End(xlUp).Row
        'LastRowFS = first_sheet.Range("I" & Rows.Count).End(xlUp).Row + 1
        'For j = 2 To LastRowI
            'first_sheet.Range("I" & LastRowFS).Value = ws.Range("I" & j)
            'first_sheet.Range("L" & LastRowFS).Value = ws.Range("L" & j)
            'LastRowFS = LastRowFS + 1
        'Next j
    'End If







