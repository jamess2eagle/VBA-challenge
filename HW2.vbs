Sub test()

Dim i, j As Long
Dim tempticker As String
Dim tempdate, volume As Double
Dim tempopen, tempclose As Double
Dim ticker(3300) As String
Dim check As Boolean
Dim yrchange, initialyr, finalyr As Double
Dim ws As Worksheet



For Each ws In Worksheets
    
    i = 2
    j = 1
    
    'labels
    ws.Range("I1") = "Ticker"
    ws.Range("j1") = "Yearly Change"
    ws.Range("k1") = "Percent Change"
    ws.Range("l1") = "Total Stock Volume"
    
    'set the first item as temp ticker
    tempticker = ws.Cells(i, 1).Value
    'put the first item in an array
    ticker(0) = tempticker
    'update the table
    ws.Cells(2, 9).Value = tempticker
    'set the initial yrchange
    initialyr = ws.Cells(i, 3).Value
    'set the initial volume
    volume = ws.Cells(i, 7).Value
    
    'while loop until end of the file
    While IsEmpty(ws.Cells(i - 1, 1)) = False
    
        'sum volume
        volume = volume + ws.Cells(i, 7).Value
        'set the first column as the temp ticker
        tempticker = ws.Cells(i, 1).Value
        
        'set initial check as false
        check = False
        
        'loop through each element in ticker
        For Each element In ticker
            
            'if tempticker is already in the array, change the check to True
            If tempticker = element Then
                check = True
            End If
        Next
        
        'if the check is false (if the element is not in the array), then update the table
        If check = False Or (IsEmpty(ws.Cells(i, 1)) = True) Then
            finalyr = ws.Cells(i - 1, 6)
            ws.Cells(j + 2, 9).Value = tempticker
            'calculate yearly chanhge
            yrchange = finalyr - initialyr
            
            
            
            'update the excel chart
            ws.Cells(j + 1, 10).Value = yrchange
            ws.Cells(j + 1, 10).NumberFormat = "0.00"
            If ws.Cells(j + 1, 10).Value > 0 Then
                ws.Cells(j + 1, 10).Interior.ColorIndex = 4
            ElseIf ws.Cells(j + 1, 10).Value < 0 Then
                ws.Cells(j + 1, 10).Interior.ColorIndex = 3
            End If
            
            'checks initial value = 0
            If initialyr = 0 Then
                ws.Cells(j + 1, 11).Value = "0"
                ws.Cells(j + 1, 11).NumberFormat = "0.00%"
            Else
            ws.Cells(j + 1, 11).Value = ws.Cells(j + 1, 10).Value / initialyr
            ws.Cells(j + 1, 11).NumberFormat = "0.00%"
            ws.Cells(j + 1, 12).Value = volume
            End If
            If volume = 0 Then
                ws.Cells(j + 1, 12).Value = "0"
            End If
            
            'set initial year
            initialyr = ws.Cells(i, 3).Value
            
            'update the array
            ticker(j) = tempticker
            j = j + 1
            'resets volume
            volume = ws.Cells(i, 7).Value
        End If
        
        i = i + 1
    
    Wend

Next


End Sub
