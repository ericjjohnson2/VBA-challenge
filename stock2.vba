Attribute VB_Name = "Module1"
Sub stock2()
    'make this work for all workbooks
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
            
            'set headers for output table1 and output table2
            Range("i1").Value = "Ticker"
            Range("j1").Value = "Yearly Change"
            Range("k1").Value = "Percent Change"
            Range("l1").Value = "Total Stock Volume"
            Range("P1").Value = "Ticker"
            Range("Q1").Value = "Value"
            Range("o2").Value = "Greatest % Increase"
            Range("o3").Value = "Greatest % Decrease"
            Range("o4").Value = "Greatest Total Volume"
            
    
            'declare some variables
            
            'i will count the rows in the data
            Dim i As Double
            
            'x will count the rows in the output table1
            Dim x As Double
            
            'x must start at 2 for the output table1
            x = 2
            
            'declare variables for keeping track of the stock open value, close value, and volume
            Dim vol As Double
            Dim opn As Double
            Dim cld As Double
            
            'ensure these variables start at zero
            vol = 0
            opn = 0
            cld = 0
            
            'create variable for the last row
            Dim LastRow As Double
            
            'determine the last row
            LastRow = Cells(Rows.Count, 1).End(xlUp).Row
            
            'create first loop to cycle through table looking for tickers and values associated with them
            
            For i = 2 To LastRow
            
                'before the if statements, we want to add the stock value; it must add for every line within the ticker
                vol = vol + Cells(i, 7).Value
                
                'first if statement is going to look for the first row of a new ticker, pull the opening value, and write to the output table1
                
                'it will check if the ticker value in the current cell is different than the one above
                If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
                        
                    'need to pull opening value and pass to variable
                    opn = Cells(i, 3).Value
                    
    
                    
                    'write ticker value to the output table- this should only happen on the first row of each ticker
                    Cells(x, 9).Value = Cells(i, 1).Value
                    
                    
                    
                End If
                
                'second if statement needs to look for the last row, take the close value, and write to the output table1.
                
                'it will check if the ticker value in the current cell is different than the one below
                If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
                    
                    'need to pull closed value
                    cld = Cells(i, 6).Value
                    
    
                    
                    
                    'write to the output table1
                    
                    'calculate the yearly change by subtracting the open value from the closed value and write to output table1
                    Cells(x, 10).Value = cld - opn
                    
                    'calculate the percentage change and write to output table1
                    Cells(x, 11).Value = (cld - opn) / opn
                    
                    'format as percentage
                    Cells(x, 11).NumberFormat = "0.00%"
                    
                    'Write volume write to output table1
                    Cells(x, 12).Value = vol
                
                    'need to incriment the ticker row before moving on
                    x = x + 1
                    
                    'Before looping to the next i, we want to set vol back to zero
                     vol = 0
                
                End If
                
                
                
            
                
                
            Next i
            
            
         
    
        'create conditional formatting for the yearly change in output table1
        
        'declare variables
        Dim YR As Range
        Dim LastRow2 As Long
        
        ' Determine the last row in column J (Yearly Change) in the output table1
        LastRow2 = Cells(Rows.Count, "J").End(xlUp).Row
        
        ' Define the range for the "Yearly Change" column
        Set YR = Range("J2:J" & LastRow2)
        
        ' Delete any existing formatting
        YR.FormatConditions.Delete
        
        ' Loop through each cell in the defined range
        For Each cell In YR
            ' Check if the cell value is not blank
            If Not IsEmpty(cell.Value) Then
                ' Apply Conditional formatting
                If cell.Value < 0 Then
                    cell.Interior.ColorIndex = 3 ' Red color for negative values
                Else
                    cell.Interior.ColorIndex = 4 ' Green color for non-negative values
                End If
            End If
        Next cell
            
            
        'now I need to check which stock had the greatest increase in volume within the output table
        
        'i2 will keep track of the rows as we cylce through the first output table
        Dim i2 As Double
        i2 = 2
        
        'DPI is the greatest percentage increase
        Dim GPI As Double
        GPI = 0
        
        'Need a string to catch the ticker name in
        Dim tick As String
        
        'Eventually we'll need one for greatest percentage decrease too
        Dim GPD As Double
        'Set GPD at a high number to ensure that the next number it checked would always be less. 
        GPD = 5000000
        
        'Eventually we'll also need GTV for greatest total volume
        Dim GTV As Double
        
        'create for loop for finding the greatest percentage increase and storing it in a value
        
        For i2 = 2 To LastRow
            If Cells(i2, 11).Value = 2 Then
                GPI = Cells(i2, 11).Value
                tick = Cells(i2, 9).Value
            End If
            
            If Cells(i2, 11).Value > GPI Then
                GPI = Cells(i2, 11).Value
                tick = Cells(i2, 9).Value
            End If
                
        
        Next i2
        
        'print GPI to the cell
        Cells(2, 16).Value = tick
        Cells(2, 17).Value = GPI
        Cells(2, 17).NumberFormat = "0.00%"
        
        'reset i2
        i2 = 2
        
        
        'create for loop for finding the greatest percentage decrease and storing it in a value
        
        For i2 = 2 To LastRow
            If Cells(i2, 11).Value = 2 Then
                GPD = Cells(i2, 11).Value
                tick = Cells(i2, 9).Value
            End If
            
            
            If Cells(i2, 11).Value < GPD Then
            
                tick = Cells(i2, 9).Value
                GPD = Cells(i2, 11).Value
            
                
            End If
                
        
        Next i2
        
        'print GPD to the cell
        Cells(3, 16).Value = tick
        Cells(3, 17).Value = GPD
        Cells(3, 17).NumberFormat = "0.00%"
        
        'reset i2
        i2 = 2
        
        'create loop to find greatest GTV
        For i2 = 2 To LastRow
        
            If Cells(i2, 12).Value = 2 Then
                GTV = Cells(i2, 12).Value
                tick = Cells(i2, 9).Value
            End If
            
            If Cells(i2, 12).Value > GTV Then
                GTV = Cells(i2, 12).Value
                tick = Cells(i2, 9).Value
            End If
    
        Next i2
            
       'print GTV to the cell
        Cells(4, 16).Value = tick
        Cells(4, 17).Value = GTV
        
    Next ws
    
End Sub


