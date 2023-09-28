Sub alphabetical_testing():

Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call RunCode
    Next
    Application.ScreenUpdating = True

End Sub

Sub RunCode()

    Dim ticker As String
    Dim y_change, p_change, g_increase, g_decrease As Double
    Dim sumary_tab_row, j, n, a As Integer
    Dim total_Svolume, g_Tvolume As LongLong
    Dim rowA, rowI As Long
           
        sumary_tab_row = 2
        j = 2
        total_Svolume = 0
        rowA = Cells(Rows.Count, 1).End(xlUp).Row


        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        Cells(1, 15).Value = "Ticker"
        Cells(1, 16).Value = "Value"
        Cells(2, 14).Value = "Greatest%Increase"
        Cells(3, 14).Value = "Greatest%Decrease"
        Cells(4, 14).Value = "Greatest Total Volume"
        
            For n = 2 To rowA
        
                If Cells(n + 1, 1).Value <> Cells(n, 1).Value Then
        
                    ticker = Cells(n, 1).Value
                    y_change = Cells(n, 6).Value - Cells(j, 3).Value
                    p_change = (y_change / Cells(j, 3).Value)
                    total_Svolume = total_Svolume + Cells(n, 7).Value

                    Range("I" & sumary_tab_row).Value = ticker
                    Range("J" & sumary_tab_row).Value = y_change
                    Range("k" & sumary_tab_row).Value = p_change
                    Range("l" & sumary_tab_row).Value = total_Svolume
                    Range("k" & sumary_tab_row).Value = Format(p_change, "Percent")
                              
                        If y_change < 0 Then
                    
                        Range("J" & sumary_tab_row).Interior.ColorIndex = 3
                    
                        Else
                    
                        Range("J" & sumary_tab_row).Interior.ColorIndex = 4
                    
                        End If

                    j = n + 1
                    total_Svolume = 0
                    sumary_tab_row = sumary_tab_row + 1

                Else

                total_Svolume = total_Svolume + Cells(n, 7).Value
        
                End If

            Next n
        
        Range("P2").Select
        ActiveCell.FormulaR1C1 = "=MAX(C[-5])"
               
        Range("P3").Select
        ActiveCell.FormulaR1C1 = "=MIN(C[-5])"
        
        Range("P4").Select
        ActiveCell.FormulaR1C1 = "=MAX(C[-4])"
        
        g_increase = Range("P2").Value
        g_decrease = Range("P3").Value
        g_Tvolume = Range("P4").Value
        
        Range("P2").Value = Format(g_increase, "Percent")
        Range("P3").Value = Format(g_decrease, "Percent")

        Range("P4").Select
        ActiveCell.FormulaR1C1 = "=MAX(C[-4])"
            
        rowI = Cells(Rows.Count, 9).End(xlUp).Row
        
            For a = 1 To rowI
            ticker = Cells(a, 9).Value
                                                   
                If Cells(a, 11).Value <> g_increase Then
                ticker = Cells(a, 9).Value
                       
                Else
                
                Cells(2, 15).Value = ticker
                          
                End If
                                                
                If Cells(a, 11).Value <> g_decrease Then
                ticker = Cells(a, 9).Value
                       
                Else
                
                Cells(3, 15).Value = ticker
                          
                End If
                
                If Cells(a, 12).Value <> g_Tvolume Then
                ticker = Cells(a, 9).Value
                       
                Else
                
                Cells(4, 15).Value = ticker
                          
                End If
            
            Next a

Columns("A:Z").AutoFit

End Sub