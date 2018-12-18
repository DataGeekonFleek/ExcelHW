Attribute VB_Name = "Module1"
Sub VBS_TEST()
'Assign Variables
Dim Ticker As Double
Dim RunningTotal As Double
Dim Rowcounter As Long


' Set-up variables
Ticker = 1
RunningTotal = 0
Rowcounter = 2

' Set EndRow as the last filled value of column A
Endrow = Range("A1", Range("A1").End(xlDown)).Rows.Count

'Select A1
Range("A1").Select


' Run the overall loop that populates the Ticker and the Volume on each Worksheet
For Each ws In Worksheets

    'Run a loop that starts in the 2nd row and goes to the end of the page
    For i = 2 To Endrow
    
        ' If i +1 cell is the same as i cell, then add the value in i cell to the runnig total
        If ws.Cells(i + 1, Ticker).Value = ws.Cells(i, Ticker) Then
        RunningTotal = RunningTotal + ws.Cells(i, 7).Value
        
        'if they don't , add the last value to the Running Total then pull the ticker and runningtotal sum into  row 9 & 10
        ElseIf ws.Cells(i + 1, Ticker).Value <> ws.Cells(i, Ticker) Then
        RunningTotal = RunningTotal + ws.Cells(i, 7).Value
        
        ws.Cells(Rowcounter, 10).Value = RunningTotal
        ws.Cells(Rowcounter, 9).Value = Cells(i, 1)
        
        'Re-set the Running Total
        RunningTotal = 0
        
        'Move the rowcounter to the next row
        Rowcounter = Rowcounter + 1
        End If
        
    Next i
    'Set the Rowcounter back to 2  for the next page
    Rowcounter = 2
    
    Next ws
                
        
End Sub


