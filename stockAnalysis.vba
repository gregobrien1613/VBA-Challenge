VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stockAnalysis()

' Declare Variables
Dim lastRow As Long, summaryTableRow As Integer, stock As String, volume As LongLong, openPrice As Long, closePrice As Long, greatestInc As Double, greatestDec As Double, greatestVol, wsCount As Integer

' Allows code to move through whole spreadsheet
wsCount = ActiveWorkbook.Worksheets.Count

For j = 1 To wsCount

    ' Places Headers
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Volume"
    Cells(1, 11).Value = "Price Change"
    Cells(1, 12).Value = "Percent Change"
    Cells(2, 14).Value = "Greatest % Increase"
    Cells(3, 14).Value = "Greatest % Decrease"
    Cells(4, 14).Value = "Greatest Volume"
    
    ' Gets Row Count for For Loop
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Allows summary table to update row by row
    summaryTableRow = 2
    
    'Declares values outside of for loop
    greatestInc = 0
    greatestDec = 0
    greatestVol = 0
    
    ' Loops through table
    For i = 2 To lastRow
    
        ' If stock ticker is not the same as the row before
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
            ' Variable for ticker
            stock = Cells(i, 1).Value
            
            ' Variable for volume
            volume = volume + Cells(i, 7).Value
            
            ' Places volume in summary table
            Range("J" & summaryTableRow).Value = Format(volume, "#,##0")
            
            ' Places ticker in summary table
            Range("I" & summaryTableRow).Value = stock
            
            ' Puts down difference between opening price and closing price
            Cells(summaryTableRow, 11).Value = Cells(i, 6).Value - Cells(i - 259, 3).Value
            
                ' Makes conditional formatting
            If Cells(summaryTableRow, 11).Value < 0 Then
                     
                ' Makes negative values red
                Cells(summaryTableRow, 11).Interior.ColorIndex = 3
                     
            Else
                     
               ' Makes positive values green
                Cells(summaryTableRow, 11).Interior.ColorIndex = 4
                
            End If
           
            If Cells(i - 259, 3).Value = 0 Then
                
                'Avoids Division by zero error
                Cells(summaryTableRow, 12).Value = Null
            
            Else
            
            'Creates Percentages
            Cells(summaryTableRow, 12).Value = Format((Cells(i, 6).Value - Cells(i - 259, 3).Value) / Cells(i - 259, 3).Value, "Percent")
            
            End If
            
            If Cells(summaryTableRow, 12).Value > greatestInc Then
                
                ' Places Greatest % Increase
                Cells(2, 15).Value = stock
                Cells(2, 16).Value = Format(Cells(summaryTableRow, 12).Value, "Percent")
            
            End If
            
            If Cells(summaryTableRow, 12).Value < greatestDec Then
            
                ' Places Greatest % Decrease
                Cells(3, 15).Value = stock
                Cells(3, 16).Value = Format(Cells(summaryTableRow, 12).Value, "Percent")
                
           End If
           
           If Range("J" & summaryTableRow).Value > greatestVol Then
            
                ' Places Largest Volume
                Cells(4, 15).Value = stock
                Cells(4, 16).Value = Format(Range("J" & summaryTableRow).Value, "Standard")
           
           End If
            
            'Updates summary table to go to next row
            summaryTableRow = summaryTableRow + 1
            
        Else
               
            ' If ticker is the same as row before add volume to summary table
            volume = volume + Cells(i, 7).Value
        
        End If
    
    Next i
    
    ' Cancels out of range error
    If j = wsCount Then
    
    Else
    
        ' Selects next sheet
         Worksheets(j + 1).Activate
     
     End If
     
Next j

End Sub

