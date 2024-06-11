
Sub quarter()

Dim first As Double
Dim last As Double
Dim counter As Integer
Dim num As Integer
Dim total As Double

    ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
    For Each ws In Worksheets
       ws.Range("I1").Value = "Ticket"
       ws.Range("J1").Value = "Quartly Change"
       ws.Range("K1").Value = "Percent Change"
       ws.Range("L1").Value = "Total Stock Volume"
        
        ' Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        counter = 1
        num = 0
        total = 0
        For i = 2 To LastRow
           
           If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
           num = num + 1
           first = ws.Cells(i - num + 1, 3).Value
           total = total + ws.Cells(i, 7).Value
           Else
           counter = counter + 1
           ws.Cells(counter, 9).Value = ws.Cells(i, 1).Value
           last = ws.Cells(i, 6).Value
           ws.Cells(counter, 10).Value = last - first
           ws.Cells(counter, 11).Value = (last - first) / first
           ws.Cells(counter, 11).NumberFormat = "0.00%"
           ws.Cells(counter, 12).Value = total + ws.Cells(i, 7).Value
           num = 0
           total = 0
           End If
        Next i
        
        For j = 2 To LastRow
            
            If ws.Cells(j, 10).Value > 0 Then
              ws.Cells(j, 10).Interior.ColorIndex = 4
            ElseIf ws.Cells(j, 10) < 0 Then
              ws.Cells(j, 10).Interior.ColorIndex = 3
             
            End If

        Next j
        
        'Calculation of "Greatest % increase", "Greatest % decrease", and "Greatest total volume"
         ws.Cells(1, 16).Value = "Ticket"
         ws.Cells(1, 17).Value = "Value"
         
         ws.Cells(2, 15).Value = "Greatest % increase"
         ws.Cells(2, 16).Value = "MSE"
         ws.Cells(2, 17).Value = Application.WorksheetFunction.Max(ws.Range("K2", "K" & LastRow))
         ws.Cells(2, 17).NumberFormat = "0.00%"
         
         ws.Cells(3, 15).Value = "Greatest % decrease"
         ws.Cells(3, 16).Value = "VNG"
         ws.Cells(3, 17).Value = Application.WorksheetFunction.Min(ws.Range("K2", "K" & LastRow))
         ws.Cells(3, 17).NumberFormat = "0.00%"
         
         ws.Cells(4, 15).Value = "Greatest total volume"
         ws.Cells(4, 16).Value = "HK"
         ws.Cells(4, 17).Value = Application.WorksheetFunction.Max(ws.Range("L2", "L" & LastRow))
    Next ws
End Sub

