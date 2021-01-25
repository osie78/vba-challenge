Attribute VB_Name = "Module1"
Sub grouping()


For Each ws In Worksheets



    Dim vol As Double
    Dim ticker As String

    Dim openprice As Double
    Dim closeprice As Double
    Dim Delta As Double
    Dim perchange As Double

    Dim counter As Integer


    'Headers and headers format
    ws.Range("j1") = "Ticker"
    ws.Range("j1").Font.FontStyle = "Bold"

    ws.Range("K1") = "Total Stock Vol"
    ws.Range("K1").Font.FontStyle = "Bold"
    ws.Columns(11).AutoFit

    ws.Range("l1") = "Yearly Change"
    ws.Range("l1").Font.FontStyle = "Bold"
    ws.Columns(12).AutoFit

    ws.Range("m1") = "% Change"
    ws.Range("m1").Font.FontStyle = "Bold"


    'ID when the stocks ticker changes
        
            'auxiliary counter for the calculation block for Open Price
            c = 0
            'Counter to build a table
            k = 2
        
    'read from row 2 to lastrow
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
   
        For i = 2 To lastrow

        ticker = ws.Cells(i, 1).Value

   
            'Verifying when the stock ticker changes in the data set
            If ticker = ws.Cells(1 + i, 1).Value Then
        
             'CALCULATION BLOCK 1: Stocks volume and determining open price
             '----------
             'Calculating the total volume of the stocks moved in the year
             vol = vol + ws.Cells(i, 7).Value
         
            'counter to keep fixed the open value of a ticker
            c = c + 1
            o = i - c
            openprice = ws.Cells(o + 1, 3).Value
            
                'addressing dividing by zero in case stock open price is 0
                If openprice = 0 Then
                    openprice = 1
                  
                Else
                End If
            
            '--------------
            'END OF CALCULATION BLOCK 1
        
            Else
            c = 0
        
        'Building the table
        
            vol = vol + ws.Cells(i, 7).Value 'volume calculation
                
            ws.Cells(k, 10).Value = ticker 'Prints tickers in consecutive rows
            ws.Cells(k, 11).Value = vol 'Prints volume in consecutive rows
             
            vol = 0 'restart variable "vol" for each loop
    
            'CALCULATION BLOCK 2
            '----------------------
            'Delta Calculation
   
            closeprice = ws.Cells(i, 6)
            If closeprice = 0 Then
                    closeprice = 1
                  
                Else
                End If
        
            Delta = closeprice - openprice
        
            '% Change calculation
            perchange = (closeprice - openprice) / openprice
            
            '----------------------
            
            'Conditional format for positive or negative price change
             
            ws.Cells(k, 12).Value = Delta
             
             
             If Delta <= 0 Then
             ws.Cells(k, 12).Interior.ColorIndex = 3
             Else
             ws.Cells(k, 12).Interior.ColorIndex = 4
             End If
        
            'Printing % change
            
            ws.Cells(k, 13).Value = perchange
            ws.Cells(k, 13).NumberFormat = "0.00%"
           
           
            k = k + 1
         
        End If

    Next i


    
Next ws
  
 
End Sub



