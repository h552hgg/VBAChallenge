Sub StockRunner()

    Dim cws As Worksheet
    Dim charttable As Boolean
    Dim chart0table As Boolean
    
    chart0table = True
    
    charttable = False
    
    For Each cws In Worksheets
    
        Dim tickervalue As String
            tickervalue = " "
        Dim ticker As String 'ticker value
            ticker = " "
        Dim totalticker As Double
            totalticker = 0
        Dim i As Long
        'Dim j As Long
        Dim lastrow As Double
        Dim totalstck As Double ' evaluate total stock value
            totalstck = 0
        Dim table As Integer 'use to establish the table consistently
            table = 2
        Dim pchange As Double ' use to calculate the change
            pchange = 0
        Dim highpchange As Double
            highpchange = 0
        Dim lowpchange As Double
            lowpchange = 0
        Dim changey As Double
            changey = 0
        Dim highvolticker As String
            highvolticker = " "
        Dim lowvolticker As String
            lowvolticker = " "
        Dim greatestticker As String
            greatestticker = " "
        
        Dim highvol As Double
            highvol = 0
        'Use theses buckets to find value
        Dim op As Double 'open price
            op = 0
        
        Dim cl As Double ' close price
            cl = 0
        
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row
        
        'lastcol = Cells(Rows.Count, "A").End(xlUp).Row
        
        If chart0table Then
        
            cws.Range("J1").Value = "Ticker"
            
            cws.Range("K1").Value = "Year Change"
            
            cws.Range("L1").Value = "Percent Change"
            
            cws.Range("M1").Value = "Total Stock"
            
            cws.Range("O2").Value = "Greatest % Increase"
            
            cws.Range("O3").Value = "Greatest % Decrease"
            
            cws.Range("O4").Value = "Greatest Total Volume"
            
            cws.Range("P1").Value = "Ticker"
            
            cws.Range("Q1").Value = "Value"
            
        
        Else
        
            charttable = True
        
        End If
        
            op = cws.Cells(2, 3).Value
        
        For i = 2 To lastrow
        
        
            If cws.Cells(i + 1, 1).Value <> cws.Cells(i, 1).Value Then
                  ticker = cws.Cells(i, 1).Value
                
                  totalstck = totalstck + cws.Cells(i, 7).Value
                
                  cl = cws.Cells(i, 6).Value
                
                  changey = cl - op
              
                    If op <> 0 Then
                       pchange = (changey / op) * 100
                    Else
                      MsgBox ("Error at ticker: " & ticker & ", with an opening price of  " & op)
                    
                    End If
                    
                      totalstck = totalstck + cws.Cells(i, 7).Value
                    
                      cws.Range("J" & table).Value = ticker
                      cws.Range("L" & table).Value = changey
           
                    
                    'If cws.Cells(i, 11).Value > 0 Then
                      'cws.Cells(i, 11).Interior.ColorIndex = 4
                    
                    'ElseIf cws.Cells(i, 11).Value < 0 Then
                      'cws.Cells(i, 11).Interior.ColorIndex = 3
                    
                    
                    
                    'End If
                    
                      cws.Range("K" & table).Value = pchange
                      cws.Range("M" & table).Value = totalstck
                    
                      table = table + 1
                      pchange = 0
                      totalstck = 0
                      
             
                    
                      op = cws.Cells(i, 3).Value
                      
                    
                    
                    If (cws.Cells(i, 12).Value > highpchange) Then
                        highpchange = cws.Cells(i, 12).Value
                        highvolticker = cws.Cells(i, 10).Value
                    
                    
                    ElseIf (cws.Cells(i, 12).Value < lowpchange) Then
                       lowpchange = cws.Cells(i, 12).Value
                       lowvolticker = cws.Cells(i, 10).Value
                    
                    
                    ElseIf (cws.Cells(i, 13).Value > totalstock) Then
                       totalticker = cws.Cells(i, 13).Value
                       greatestticker = cws.Cells(i, 10).Value
                       
                    End If
                  cws.Range("P4").Value = greatestticker
                  cws.Range("P2").Value = highvolticker
                  cws.Range("P3").Value = lowvolticker
                  cws.Range("Q2").Value = (CStr(highpchange) & "%")
                  cws.Range("Q3").Value = (CStr(lowpchange) & "%")
                  cws.Range("Q4").Value = totalticker
                 
        
          
                   pchange = 0
                   totalticker = 0
          
            Else
              totalticker = totalticker + cws.Cells(i, 7).Value
                    
                
            End If
    
        
        Next i
        For i = 2 To lastrow
                   If cws.Cells(i, 11).Value > 0 Then
                      cws.Cells(i, 11).Interior.ColorIndex = 4
                    
                    ElseIf cws.Cells(i, 11).Value < 0 Then
                      cws.Cells(i, 11).Interior.ColorIndex = 3
                    
                    
                    
                    
                    End If
      Next i
        
           
        
        
        
        
        
        
    
    
    
    Next cws
    
      

End Sub