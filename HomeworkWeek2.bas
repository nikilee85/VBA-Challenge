Attribute VB_Name = "Module1"
Sub Stocking()
summarycount = 2

    ticker = 0
    firstCheck = True
    firstOpen = 0
    lastClose = 0
    totalvolume = 0
    yearlyChange = 0
    
    Dim lastrow As Long
    lastrow = ActiveSheet.UsedRange.Rows.Count
    For I = 2 To lastrow
    
    If firstCheck = True Then
        firstCheck = False
        firstOpen = Cells(I, 3).Value
    
    End If
    
    If Cells(I + 1, "A").Value = Cells(I, "A") Then
     totalvolume = totalvolume + Cells(I, "G").Value
    
    Else
        
        'ticker counter
        ticker = Cells(I, 1).Value
        Cells(summarycount, "I").Value = ticker
       
       'total volume
        totalvolume = totalvolume + Cells(I, "G").Value
        Cells(summarycount, "L").Value = totalvolume
        
        'delta year
        lastClose = Cells(I, "F").Value
        yearlyChange = lastClose - firstOpen
        Cells(summarycount, "J").Value = yearlyChange
        
        'delta percent
        
        If (firstOpen + lastClose >= 0) Then
        percentChange = (yearlyChange / firstOpen) * 100
        Cells(summarycount, "K").Value = percentChange
        
        Else
        Cells(summarycount, "K").Value = 0
        
        End If
        
        'change cell colour
        If percentChange > 0 Then Cells(summarycount, "K").Interior.ColorIndex = 4
        If percentChange < 0 Then Cells(summarycount, "K").Interior.ColorIndex = 3
        
        'next row counter
        firstCheck = True
        totalvolume = 0
        firstOpen = 0
        lastClose = 0
        yearlyChange = 0
        summarycount = summarycount + 1
End If
Next I

End Sub

Sub Clear()
   Dim lastrow As Long
    lastrow = ActiveSheet.UsedRange.Rows.Count
    
Range("I2:L" & lastrow) = ""
Range("K2:K" & lastrow).Interior.ColorIndex = 0

End Sub

