Sub test()

Dim i, j As Long
Dim tempticker As String
Dim tempdate, volume As Double
Dim tempopen, tempclose As Double
Dim ticker(3300) As String
Dim check As Boolean
Dim yrchange, initialyr, finalyr As Double
Dim ws As Worksheet
Dim highper, lowper, highvol As Long
Dim temphighper, templowper, temphighvol As Double




For Each ws In Worksheets
    
    i = 2
    j = 1
    
    'labels
    ws.Range("I1") = "Ticker"
    ws.Range("j1") = "Yearly Change"
    ws.Range("k1") = "Percent Change"
    ws.Range("l1") = "Total Stock Volume"
    ws.Range("n2") = "greatest % increase"
    ws.Range("n3") = "greatest % decrease"
    ws.Range("n4") = "greatest total volume"
    ws.Range("o1") = "Ticker"
    ws.Range("p1") = "value"

    
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
    
    i = 2
    highper = 2
    lowper = 2
    highvol = 2
    
    'set temp = first value
    temphighper = ws.Cells(2, 11).Value
    templowper = ws.Cells(2, 11).Value
    temphighvol = ws.Cells(2, 12).Value
    
    'loop until empty
    While IsEmpty(ws.Cells(i - 1, 1)) = False
        'if there is a higher percentage, replace the value and store row number
        If temphighper < ws.Cells(i, 11).Value Then
            temphighper = ws.Cells(i, 11).Value
            highper = i
        End If
        If templowper > ws.Cells(i, 11).Value Then
            templowper = ws.Cells(i, 11).Value
            lowper = i
        End If
        If temphighvol < ws.Cells(i, 12).Value Then
            temphighvol = ws.Cells(i, 12).Value
            highvol = i
        End If
    
        i = i + 1
    Wend
    
    'update excel
    ws.Cells(2, 15).Value = ws.Cells(highper, 9).Value
    ws.Cells(2, 16).Value = ws.Cells(highper, 11).Value
    ws.Cells(2, 16).NumberFormat = "0.00%"
    ws.Cells(3, 15).Value = ws.Cells(lowper, 9).Value
    ws.Cells(3, 16).Value = ws.Cells(lowper, 11).Value
    ws.Cells(3, 16).NumberFormat = "0.00%"
    ws.Cells(4, 15).Value = ws.Cells(highvol, 9).Value
    ws.Cells(4, 16).Value = ws.Cells(highvol, 12).Value
    
    
Next


End Sub
