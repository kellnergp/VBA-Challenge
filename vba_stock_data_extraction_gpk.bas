Attribute VB_Name = "Module1"
Sub dataextraction():

'bonus: run through all sheets on a loop
Dim sheetnum As Integer
sheetnum = ActiveWorkbook.Sheets.Count

For k = 1 To sheetnum

    'NOTE: It took me about 7-8 minutes per sheet for this version of the code to run
    
    'activate each sheet as the loop progresses, the rest of the code will run on each one
    ActiveWorkbook.Sheets(k).Activate
    

     'set column headers
     Range("I1").Value = "Ticker"
     Range("J1").Value = "Yearly Change"
     Range("K1").Value = "Percent Change"
     Range("L1").Value = "Total Stock Volume"
     
     'define variables for collating unique tickers
     Dim tickcounter As Integer
     Dim currenttick As String
     
     'grab first value as guaranteed unique ticker
     Range("I2") = Cells(2, 1)
     
     'set counter to start at first cell for processed data
     tickcounter = 2
     
     'determine length of data set
     Dim lastrow As Long
     lastrow = ActiveSheet.UsedRange.Rows.Count
     
     'loop through whole dataset to find all unique tickers
     'worksheetfunction.unique would also work but it isn't in all versions of excel
     For i = 2 To lastrow
         currenttick = Cells(i, 1)
         If currenttick <> Cells(tickcounter, 9) Then
             tickcounter = tickcounter + 1
             Cells(tickcounter, 9) = currenttick
         End If
     Next i
     
    'get a count of all unique tickers
    Dim cellcount As Integer
    cellcount = WorksheetFunction.CountA(Range("I:I"))
    
    'pull the data into an array to save time during loops
    Dim wholedata() As Variant
    ReDim wholedata(lastrow, 6)
    
    wholedata = Range(Cells(1, 1), Cells(lastrow, 7)).Value
    
    
    'pull unique tickers into an array for reference
    Dim uniquetickers As Variant
    ReDim uniquetickers((cellcount - 1))
    uniquetickers = Range(Cells(2, 9), Cells(cellcount, 9))
    
    'Range("J2:J2836").Value = uniquetickers
    
    'declare variables for data storage
    Dim startopen() As Variant
    ReDim startopen(0 To (cellcount - 1)) As Variant
    Dim endclose() As Variant
    ReDim endclose(0 To (cellcount - 1)) As Variant
    Dim runningvolume() As Variant
    ReDim runningvolume(0 To (cellcount - 1)) As Variant
    
    
    'loop through data, using conditionals and an iterating counter to generate arrays of pertinent data for each ticker
    For i = 1 To (cellcount - 1)
    
        
        runningvolume(i) = 0 'set an initial value to make things easier inside nested loop
        
        For j = 2 To (lastrow - 1)
            Dim currentticker As String
            currentticker = uniquetickers(i, 1)
            Dim dataticker As String
            dataticker = wholedata(j, 1)
            Dim prevticker As String
            prevticker = wholedata(j - 1, 1)
            
            'if current row has correct ticker but previous row doesn't its the first row of that ticker
            If (dataticker = currentticker) And (prevticker <> currentticker) Then
                'in that case, pull the starting value for opening price and add this row to the running volume total
                startopen(i) = wholedata(j, 3)
                runningvolume(i) = runningvolume(i) + wholedata(j, 7)
                
            'the rest of the rows that have the current ticker
            ElseIf dataticker = currentticker Then
                'add to the running volume total and pull this rows closing value so that the last value pulled will naturally be the end of year close
                runningvolume(i) = runningvolume(i) + wholedata(j, 7)
                endclose(i) = wholedata(j, 6)
            
            'previous row had correct ticker, current row doesn't means nested loop is done, move on to next ticker
            ElseIf dataticker <> currentticker And prevticker = currentticker Then
                Exit For
            End If
        
        Next j
    
    Next i
            
    For i = 1 To (cellcount - 1)
        
        'yearly change is final close minus starting open
        Cells(i + 1, 10) = endclose(i) - startopen(i)
    
        'use a conditional to account for 0 values
        If endclose(i) = 0 And startopen(i) = 0 Then
            Cells(i + 1, 11) = 0
        ElseIf startopen(i) = 0 Then
            Cells(i + 1, 11) = "N/A" 'infinite value isn't useful
        Else
            Cells(i + 1, 11) = (endclose(i) - startopen(i)) / startopen(i)
        End If
        
        'total stock volume is the final value from the running total
        Cells(i + 1, 12) = runningvolume(i)
        
    Next i
     
    'set format for percentage column
    Range(Cells(2, 11), Cells(cellcount, 11)).NumberFormat = "0.00%"
    
    'conditional color format for yearly change column
    For i = 2 To cellcount
        
        'green for growth, red for loss
        If Cells(i, 10) > 0 Then
            Cells(i, 10).Interior.ColorIndex = 4
        ElseIf Cells(i, 10) < 0 Then
            Cells(i, 10).Interior.ColorIndex = 3
        End If
    
    Next i
     
Next k
    
End Sub
