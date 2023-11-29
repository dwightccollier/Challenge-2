Attribute VB_Name = "Module1"



Option Explicit

Sub stock()

Dim FTable As Worksheet

For Each FTable In Worksheets

    
    
    'Create a variable to hold the ticker symbol, date, total volume and percent
    

   
   Dim Ticker As String
   
   
   
   Dim TickerCount As Double
   TickerCount = 0
   
   Dim Lastrow As Long
   Dim i As Long
   

   

        Lastrow = FTable.Cells(Rows.Count, 1).End(xlUp).Row
        FTable.Range("I1").EntireColumn.Insert
        FTable.Range("J1").EntireColumn.Insert
        FTable.Range("K1").EntireColumn.Insert
        FTable.Range("L1").EntireColumn.Insert
        FTable.Cells(1, 9).Value = "Ticker"
        FTable.Cells(1, 10).Value = "Yearly Change"
        FTable.Cells(1, 11).Value = "Percent Change"
        FTable.Cells(1, 12).Value = "Total Stock Volume"
        
   

     FTable.Range("P1").Value = "Ticker"
     FTable.Range("Q1").Value = "Value"
     FTable.Range("O2").Value = "Greatest % Increase"
     FTable.Range("O3").Value = "Greatest % Decrease"
     FTable.Range("O4").Value = "Greatest Total Volume"
     
     
     
     
      Dim openprice As Double
      Dim closeprice As Double
      Dim pricechange As Double
      Dim percentpricechange As Double
      Dim Formatchange As Double
      Dim incTicker As String
      Dim incVal As Double
      Dim decTicker As String
      Dim decVal As Double
      Dim greatestTicker As String
      Dim greatestVolume As Double
      
      
      openprice = 0
      closeprice = 0
      pricechange = 0
      percentpricechange = 0
      incVal = 0
      decVal = 0
      greatestVolume = 0

    Dim TickerRow As Long: TickerRow = 1

        
   
      For i = 2 To Lastrow

   If openprice = 0 Then

          openprice = FTable.Cells(i, 3).Value
      End If
      
   
     If FTable.Cells(i + 1, 1).Value <> FTable.Cells(i, 1).Value Then
     
     
      TickerRow = TickerRow + 1
     

     Ticker = FTable.Cells(i, 1).Value
     
     FTable.Cells(TickerRow, "I").Value = Ticker
     
     '

        
          closeprice = FTable.Cells(i, 6).Value
     
     
    pricechange = closeprice - openprice
    
    
    FTable.Cells(TickerRow, "J").Value = pricechange

    If pricechange < 0 Then
            FTable.Cells(TickerRow, "J").Interior.ColorIndex = 3
        ElseIf pricechange > 0 Then
            FTable.Cells(TickerRow, "J").Interior.ColorIndex = 4


     End If
    

    

    
     
     percentpricechange = (pricechange / openprice)


     FTable.Cells(TickerRow, "K").Value = percentpricechange
     
     FTable.Cells(TickerRow, "K").NumberFormat = ".##%"
     
     If incVal < percentpricechange Then

     incVal = percentpricechange

     incTicker = Ticker

     End If

     If decVal > percentpricechange Then

     decVal = percentpricechange

     decTicker = Ticker

     End If
      
    
    openprice = 0


    
    
    TickerCount = TickerCount + FTable.Cells(i, 7).Value
    
    FTable.Cells(TickerRow, "L").Value = TickerCount
    FTable.Cells(TickerRow, "L").NumberFormat = "0"

   If greatestVolume < TickerCount Then

   greatestVolume = TickerCount

   greatestTicker = Ticker
   
   End If
   

    
    TickerCount = 0
    
    
    

    
    Else

    
    TickerCount = TickerCount + FTable.Cells(i, 7).Value

    End If
     
    Next i
    
    FTable.Cells(2, "P") = incTicker
    FTable.Cells(2, "Q") = incVal

    FTable.Range("Q1:Q3").NumberFormat = ".##%"

    FTable.Cells(3, "P") = decTicker
    FTable.Cells(3, "Q") = decVal

    FTable.Cells(4, "P") = greatestTicker
    FTable.Cells(4, "Q") = greatestVolume

    Next FTable
    
End Sub


