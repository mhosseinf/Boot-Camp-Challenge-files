Attribute VB_Name = "Module1"
Sub stocks()
    ' Define variations
    Dim Summary_table_row As Integer
    Dim Summary_table_row1 As Integer
    Dim Lastrow As Double
    Dim i As Double
    Dim j As Double
    Dim ticker_name As String
    Dim ticker_volume As Double
    Dim Greatesttotalvolum As Double
    Dim Greatestthicker As String
    Dim opendate As Double
    Dim opendatevalue As Double
    Dim closedatevalue As Double
    Dim yearlychange As Double
    Dim yearlychangepercent As Double
    Dim Greatestpercentincrease As Double
    Dim Greatestpercentdecrease As Double
    Dim Lowestthicker As String
    
    
        
    
    'initial value
    
    Summary_table_row = 2
    Summary_table_row1 = 1
    opendate = ActiveSheet.Cells(2, 2).Value
    opendatevalue = ActiveSheet.Cells(2, 3).Value
    
    'Populate column headers
    ActiveSheet.Range("K1").Value = "Ticker"
    ActiveSheet.Range("N1").Value = "Total Stock Volum"
    ActiveSheet.Range("L1").Value = "Yearly Change"
    ActiveSheet.Range("M1").Value = "Percent Change"
    ActiveSheet.Range("Q2").Value = "Greatest%Increase"
    ActiveSheet.Range("Q3").Value = "Greatest%Decrease"
    ActiveSheet.Range("Q4").Value = "GreatestTotalvolum"
    ActiveSheet.Range("R1").Value = "Thicker"
    ActiveSheet.Range("S1").Value = "Value"
    
    'Cout no of rows
    Lastrow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
    
    'loop to go through all rows one by one and summarise thicker and calculate The total stock volume of the associated stock
    For i = 2 To Lastrow
     If ActiveSheet.Cells(i + 1, 1).Value <> ActiveSheet.Cells(i, 1).Value Then
     ticker_name = ActiveSheet.Cells(i, 1).Value
     ticker_volume = ticker_volume + ActiveSheet.Cells(i, 7).Value
     ActiveSheet.Range("K" & Summary_table_row).Value = ticker_name
     ActiveSheet.Range("N" & Summary_table_row).Value = ticker_volume
     Summary_table_row = Summary_table_row + 1
     
      'steps to calculate Greatest total volume and associated thicker
        If Greatesttotalvolum < ticker_volume Then
            Greatesttotalvolum = ticker_volume
            Greatestthicker = ActiveSheet.Cells(i, 1).Value
        End If
     ActiveSheet.Range("R4").Value = Greatestthicker
     ActiveSheet.Range("S4").Value = Greatesttotalvolum
     ticker_volume = 0
     
     Else
     ticker_volume = ticker_volume + ActiveSheet.Cells(i, 7).Value
     End If
       
    Next i
    
    'loop to go through all rows one by one and summarise thicker and calculate Yearly change/ The percentage change
     For j = 2 To Lastrow
    
        If ActiveSheet.Cells(j + 1, 1).Value = ActiveSheet.Cells(j, 1).Value Then
            
            If ActiveSheet.Cells(j + 1, 2).Value > opendate Then
            closedatevalue = ActiveSheet.Cells(j + 1, 6).Value
            
            Else
            opendate = ActiveSheet.Cells(j, 2).Value
            opendatevalue = ActiveSheet.Cells(j, 3).Value
            
            End If
         Else
         Summary_table_row1 = Summary_table_row1 + 1
         yearlychange = (closedatevalue - opendatevalue)
         yearlychangepercent = (closedatevalue - opendatevalue) / opendatevalue
            
            'steps to calculate Greatest % increase and associated thicker
            If Greatestpercentincrease < yearlychangepercent Then
            Greatestpercentincrease = yearlychangepercent
            Greatestthicker = ActiveSheet.Cells(j, 1).Value
            End If
            
            'steps to calculate Greatest % Decrease and associated thicker
            If Greatestpercentdecrease > yearlychangepercent Then
            Greatestpercentdecrease = yearlychangepercent
            Lowestthicker = ActiveSheet.Cells(j, 1).Value
            End If
         ActiveSheet.Range("L" & Summary_table_row1).Value = yearlychange
         ActiveSheet.Range("M" & Summary_table_row1).Value = yearlychangepercent
         ActiveSheet.Range("S2").Value = Greatestpercentincrease
         ActiveSheet.Range("S3").Value = Greatestpercentdecrease
         ActiveSheet.Range("R2").Value = Greatestthicker
         ActiveSheet.Range("R3").Value = Lowestthicker
         yearlychange = 0
         yearlychangepercent = 0
         opendate = ActiveSheet.Cells(j, 2).Value
         opendatevalue = ActiveSheet.Cells(j, 3).Value
            
            'conditional formatting Yearly change, positive green and negative red.
            If ActiveSheet.Range("L" & Summary_table_row1).Value > 0 Then
            ActiveSheet.Range("L" & Summary_table_row1).Interior.ColorIndex = 4
            Else
            ActiveSheet.Range("L" & Summary_table_row1).Interior.ColorIndex = 3
            End If
            
            
            'conditional formatting percent change, positive green and negative red.
            If ActiveSheet.Range("M" & Summary_table_row1).Value > 0 Then
            ActiveSheet.Range("M" & Summary_table_row1).Interior.ColorIndex = 4
            Else
            ActiveSheet.Range("M" & Summary_table_row1).Interior.ColorIndex = 3
            End If
         End If
    Next j
    'Change style to percent for: Yearly %change,Greatest%Increase, Greatest% Decrease

    ActiveSheet.Range("M:M").NumberFormat = "0.00%"
    ActiveSheet.Range("S2").NumberFormat = "0.00%"
    ActiveSheet.Range("S3").NumberFormat = "0.00%"
    ActiveSheet.Range("k:s").EntireColumn.AutoFit
   
    
End Sub


