Sub vbatickerstock()



' https://www.w3schools.com/asp/asp_ref_vbscript_functions.asp
' loop through each worksheet
For Each ws In Worksheets

' variables


Dim Lastrow As Long ' found on https://www.educba.com/vba-last-row/
Dim ticker As String
Dim yropen As Double
'yropen = 0
Dim yrclose As Double
'yrclose = 0
Dim vol As Double
'vol = 0
Dim yrchange As Double
Dim perchange As Double
'perchange = CLng(10000.45)
Dim tickersum As Integer
tickersum = 1
'(hard) to find the ticker with the largest number in order to find the greatest % increase
' ticker for greatest % increase
Dim grinticker As String
' ticker for greatest % decrease
Dim grdeticker As String
' ticker for total volume
Dim grtotvolticker As String
' greatest value for greatest % increase
Dim grinvalue As Double
' greatest value for greatest % decrease
Dim grdevalue As Double
' greatest total volume
Dim grvolvalue As Double



' loop through each worksheet
'For Each ws In Worksheets


'add column names with each worksheet
ws.Cells(1, 9).Value = "Ticker Symbol"
ws.Cells(1, 10).Value = "Yearly Change ($)"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"

ws.Range("Q2").Value = 0
ws.Range("Q3").Value = 99999
ws.Range("Q4").Value = 0

'to find the last row in each worksheet
'Lastrow = Cells(Row.Count, 1).End(xlUp).Row '(did not work?!?
Lastrow = Range("A:A").SpecialCells(xlCellTypeLastCell).Row ' got this off of https://www.educba.com/vba-last-row/

Column = 1
tickersum = 1

    For i = 2 To Lastrow
    
    
    
        ' condition to find unique ticker value from the credit card class
        'tried the one for credit card but did not work
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
         
        
            'only taking last value***********************
            'find each ticker in column a
            ticker = ws.Cells(i, 1).Value
            'adds values *********NEED TO FIND OUT DOES THIS ADD UP TO GET THE DIFFERENCE FROM CLOSED??????????
            yropen = yropen + ws.Cells(i, 3).Value
            'adds values *********jsut like the yr open
            yrclose = yrclose + ws.Cells(i, 6).Value
            'adds the volume cells
            vol = vol + ws.Cells(i, 7).Value
            'find how to pull the largest and lowest values
            grvolvalue = grvolvalue + ws.Cells(i, 7).Value
            grinvalue = grinvalue + ws.Cells(i, 11).Value
            grdevalue = grdevalue + ws.Cells(i, 11).Value
            
            'going off yrchange data to set the colors
            
            
            yrchange = yrclose - yropen
            
            
            
            'If yropen <> 0 Then
            'not sure if this is 100% right
            
            perchange = (yrchange / (1 + yrclose) * 100)
            
            
            
            'from credit card
            tickersum = tickersum + 1
            ws.Range("I" & tickersum).Value = ticker
            ws.Range("J" & tickersum).Value = yrchange
            'If (yrchange >= 0) Then
            'ws.Range("J" & tickersum).Interior.ColorIndex = 4
            'ElseIf (yrchange < 0) Then
            'ws.Range("J" & tickersum).Interior.ColorIndex = 3
            ws.Range("K" & tickersum).Value = perchange
            ws.Range("L" & tickersum).Value = vol
        
            If (yrchange & tickersum >= 0) Then
                ws.Range("J" & tickersum).Interior.ColorIndex = 4
            
            ElseIf (yrchange & tickersum < 0) Then
                ws.Range("J" & tickersum).Interior.ColorIndex = 3
                
            End If
            
        
            vol = 0
            yropen = 0
            yrclose = 0
            grvolvalue = 0
            grinvalue = 0
            grdevalue = 0
            
        Else
            yropen = yropen + ws.Cells(i, 3).Value
            yrclose = yrclose + ws.Cells(i, 6).Value
            vol = vol + ws.Cells(i, 7).Value
        End If
        
    ' completes this loop for values, needs next iteration
    
    Next i

    i = 2
    Do Until ws.Cells(i, 9) = ""
    
        If ws.Cells(i, 11) > ws.Range("Q2").Value Then
        
            ws.Range("P2").Value = ws.Cells(i, 9)
            ws.Range("Q2").Value = ws.Cells(i, 11)
        End If
        
        If ws.Cells(i, 11) < ws.Range("Q3").Value Then
        
            ws.Range("P3").Value = ws.Cells(i, 9)
            ws.Range("Q3").Value = ws.Cells(i, 11)
        End If
        
        If ws.Cells(i, 12) > ws.Range("Q4").Value Then
        
            ws.Range("P4").Value = ws.Cells(i, 9)
            ws.Range("Q4").Value = ws.Cells(i, 12)
        End If
    
        i = i + 1
    Loop

Next ws
'fingers crossed!!!

End Sub



