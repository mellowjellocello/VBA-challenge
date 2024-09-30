Attribute VB_Name = "Module1"
Sub stockinformation()

Dim ws As Worksheet
Dim lastrow As Double

Dim currentticker As String
Dim quarteropen As Double
Dim quarterclose As Double
Dim quarterchange As Double
Dim totalstock As Double
Dim row As Long
Dim tickercount As Integer

Dim row2 As Long
Dim gincrease As Double
Dim gdecrease As Double
Dim gtotalvolume As Double
Dim gincreaseticker As String
Dim gdecreaseticker As String
Dim gtotalvolumeticker As String


For Each ws In Worksheets

    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).row
    
    totalstock = 0
    tickercount = 0
    quarteropen = ws.Range("C2").Value
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Quarter Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"

    For row = 2 To lastrow
        currentticker = ws.Cells(row, 1).Value
        totalstock = totalstock + ws.Cells(row, 7).Value
        
        If ws.Cells(row, 1).Value <> ws.Cells(row + 1, 1).Value And row + 1 <= lastrow Then
            tickercount = tickercount + 1
            quarterclose = ws.Cells(row, 6).Value
            quarterchange = quarterclose - quarteropen
            
            ws.Range("I" & tickercount + 1).Value = currentticker
            ws.Range("J" & tickercount + 1).Value = quarterchange
            ws.Range("K" & tickercount + 1).Value = quarterchange / quarteropen
            ws.Range("L" & tickercount + 1).Value = totalstock
            
            quarteropen = ws.Cells(row + 1, 3).Value
            totalstock = 0
            
        End If
        
    Next row


    lastrow2 = ws.Cells(Rows.Count, 9).End(xlUp).row
    
    gincrease = 0
    gdecrease = 0
    gtotalvolume = 0
    
    ws.Range("O1").Value = "Ticker"
    ws.Range("P1").Value = "Value"
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest Total Volume"
    
    For row2 = 2 To lastrow2
        If ws.Cells(row2, 11).Value > gincrease Then
            gincrease = ws.Cells(row2, 11).Value
            gincreaseticker = ws.Cells(row2, 9).Value
        ElseIf ws.Cells(row2, 11).Value < gdecrease Then
            gdecrease = ws.Cells(row2, 11).Value
            gdecreaseticker = ws.Cells(row2, 9).Value
        ElseIf ws.Cells(row2, 12).Value > gtotalvolume Then
            gtotalvolume = ws.Cells(row2, 12).Value
            gtotalvolumeticker = ws.Cells(row2, 9).Value
        
        End If
        
        ' this is the conditional formating with vba, if supposed to be done without vba,
        ' refer to stock data excel formatting.xlsm or stock data excel formatting.bas instead

        If ws.Cells(row2, 10).Value > 0 Then
            ws.Cells(row2, 10).Interior.Color = &H59BB9B
        ElseIf ws.Cells(row2, 10).Value < 0 Then
            ws.Cells(row2, 10).Interior.Color = &H4D50C0
        ElseIf ws.Cells(row2, 10).Value = 0 Then
            ws.Cells(row2, 10).Interior.Color = &H47DCE7
    
        End If
    
        If ws.Cells(row2, 11).Value > 0 Then
            ws.Cells(row2, 11).Interior.Color = &H59BB9B
        ElseIf ws.Cells(row2, 11).Value < 0 Then
            ws.Cells(row2, 11).Interior.Color = &H4D50C0
        ElseIf ws.Cells(row2, 11).Value = 0 Then
            ws.Cells(row2, 11).Interior.Color = &H47DCE7
        
        End If
        
    Next row2
            
    ws.Range("O2").Value = gincreaseticker
    ws.Range("P2").Value = gincrease
    ws.Range("O3").Value = gdecreaseticker
    ws.Range("P3").Value = gdecrease
    ws.Range("O4").Value = gtotalvolumeticker
    ws.Range("P4").Value = gtotalvolume
    
Next ws

End Sub
