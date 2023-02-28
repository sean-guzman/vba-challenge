Sub StockInfo()
    
    
    Dim ticker As String
    Dim total_volume As Double
    Dim opening As Double
    Dim closing As Double
    Dim year_change As Double
    Dim percent_change As Double
    
    Dim i, j, k As Long
        
    For Each ws In Worksheets
    
        j = 2 ' starts on second row for reporting
        k = 0 ' counter for current ticker
        
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row ' count number of rows in dataset
        
        total_volume = 0
    
        For i = 2 To lastrow
        
            ticker = ws.Cells(i, 1).Value
            total_volume = ws.Cells(i, 7).Value + total_volume
            
            ' perform following If statement if the next ticker is different
            
            If ws.Cells(i + 1, 1).Value <> ticker Then '
            
                ws.Cells(j, 9).Value = ticker ' apply ticker symbol to report
                ws.Cells(j, 12).Value = total_volume ' apply ticker's total volume to report
                
                ' perform year change and percent change calculation
                opening = ws.Cells(i - k, 3).Value
                closing = ws.Cells(i, 6).Value
                year_change = closing - opening
                percent_change = (year_change / opening)
                
                ' display year change and percent change values
                ws.Cells(j, 10).Value = year_change
                ws.Cells(j, 11).Value = percent_change
                ws.Cells(j, 11).NumberFormat = "0.00%"
                
                
                ' perform conditional formatting based on value
                If ws.Cells(j, 10).Value <= 0 Then
                
                    ws.Cells(j, 10).Interior.Color = vbRed
                    
                Else
                
                    ws.Cells(j, 10).Interior.Color = vbGreen
                    
                End If
                
                total_volume = 0 ' reset total_volume for next ticker
                
                k = 0 ' reset ticker counter
                j = j + 1 ' start next row for report
                
            Else
            
                k = k + 1 ' track number of rows of current ticker
                
                
            End If
        
        Next i


        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
    
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest & Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        lastrowofReport = ws.Cells(Rows.Count, 9).End(xlUp).Row ' count number of rows from report

        For i = 2 To lastrowofReport
            
            ' Find Greatest % Increase
            If ws.Cells(i, 11).Value > ws.Range("Q2").Value Then
            
                ws.Range("P2").Value = ws.Cells(i, 9).Value
                ws.Range("Q2").Value = ws.Cells(i, 11).Value
                ws.Range("Q2").NumberFormat = "0.00%"
            
            Else
            
            End If
        
            ' Find Greatest % Decrease
            If ws.Cells(i, 11).Value < ws.Range("Q3").Value Then
            
                ws.Range("P3").Value = ws.Cells(i, 9).Value
                ws.Range("Q3").Value = ws.Cells(i, 11).Value
                ws.Range("Q3").NumberFormat = "0.00%"
            
            Else
            
            End If
        
            ' Find Greatest % Total Volume
            If ws.Cells(i, 12).Value > ws.Range("Q4").Value Then
            
                ws.Range("P4").Value = ws.Cells(i, 9).Value
                ws.Range("Q4").Value = ws.Cells(i, 12).Value
            
            Else
            
            End If
            
        Next i
    
    Next ws
    
 
End Sub
