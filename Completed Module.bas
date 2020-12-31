Attribute VB_Name = "Module1"
Sub Test_Data():

'Creating Headings

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"

'Defining Variables
Dim j As Integer
Dim Stock_Open As Double
Dim Stock_Close As Double
Dim Percent_Change As Double
Dim Total_Volume As Double

    j = 1

'probably could have used a "while" loop here"

For i = 1 To 10000000

'Exit Statement when there is no Ticker (so this wont take an hour to run)

If Cells(i, 1).Value = "" Then

        Exit For

'Each time the ticker changes, the summary data will be populated and the nested "if" will format the color

    ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value And Cells(i, 1).Value <> "<ticker>" Then
    
        Stock_Close = Cells(i, 6).Value
        Cells(j, 9) = Cells(i, 1).Value
        Cells(j, 10) = Stock_Close - Stock_Open
        If Stock_Open <> 0 Then
        
            Cells(j, 11) = (Stock_Close / Stock_Open) - 1
            
            Else: Cells(j, 11) = ""
        
        End If
        
        Cells(j, 11).NumberFormat = "0.00%"
        Cells(j, 12) = Total_Volume
        
        If Cells(j, 10).Value >= 0 Then
            Cells(j, 10).Interior.ColorIndex = 4
            
            Else: Cells(j, 10).Interior.ColorIndex = 3
            
        End If
        
'This resets the counter on the stock volume and defines the opening price for the next stock

        Total_Volume = 0
        Stock_Open = Cells(i + 1, 3).Value
    
        j = j + 1
    
'this is just to deal with headings and begin counting the data from the first stock
    
    ElseIf Cells(i, 1).Value = "<ticker>" Then
    
        Stock_Open = Cells(i + 1, 3).Value
        
        j = j + 1
        
        Total_Volume = Cells(i + 1, 7).Value

'Whenever the ticker is the same as the previous cell, the volume will be added to the volume total

    ElseIf Cells(i, 1).Value = Cells(i + 1, 1).Value Then
    
        Total_Volume = Total_Volume + Cells(i + 1, 7).Value
        
    End If

'Bonus

Next i

For k = 2 To 10000

    If Cells(k, 11).Value = "" Then

Exit For

    ElseIf Cells(k, 11).Value > Cells(2, 17).Value Then
    
    Cells(2, 17).Value = Cells(k, 11).Value
    Cells(2, 16).Value = Cells(k, 9).Value
    
    ElseIf Cells(k, 11).Value <= Cells(3, 17).Value Then
    
    Cells(3, 17).Value = Cells(k, 11).Value
    Cells(3, 16).Value = Cells(k, 9).Value
    
End If
    
If Cells(k, 12).Value > Cells(4, 17).Value Then

    Cells(4, 17).Value = Cells(k, 12).Value
    Cells(4, 16).Value = Cells(k, 9).Value
    
End If

Next k
    
Cells(2, 17).NumberFormat = "0.00%"
Cells(3, 17).NumberFormat = "0.00%"
    
End Sub
