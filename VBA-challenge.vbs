Sub Multiple_Year_Sheet()


     ' Set variable for worksheet

     Dim ws As Worksheet
 
     ' Set summary table

     Dim Summary_Table_Header As Boolean
     Dim COMMAND_SPREADSHEET As Boolean
    
     ' Set summary table header 

     Summary_Table_Header= False 

     ' Set command for max values 

     COMMAND_SPREADSHEET = True              
    
     ' Loop through all sheets in active workbook 

    For Each ws In Worksheets


    
         ' Declare and set initial variable to hold the ticker name

         Dim Ticker As String
         Ticker = " "
        
         ' Declare and Set an initial variable for total stock volume per ticker name
         Dim total_stock_volume As Double
         total_stock_volume = 0
        
         ' Declare and Set new variables for open, close price, difference in price, percent change and yearly change

         Dim YrOp As Double
         YrOp = 0
         Dim YrCl As Double
         YrCl = 0
         Dim DfInRow As Double
         DfInRow = 0
         Dim Percent_Change As Double
         Percent_Change = 0

         ' Declare and set variables max value of ticker, min value of ticker, 
         'max percent change, min Percent Change, min total stock volume and max total stock volume

         Dim Ticker_max As String
         Ticker_max = " "
         Dim Ticker_min As String
         Ticker_min = " "
         Dim Percent_Change_max As Double
         Percent_Change_max = 0
         Dim Percent_Change_min As Double
         Percent_Change_min = 0
         Dim total_stock_volume_min_ticker As String
         total_stock_volume_min_ticker = " "
         Dim total_stock_volume_min As Double
         total_stock_volume_min = 0
         Dim total_stock_volume_max As Double
         total_stock_volume_max = 0
         Dim total_stock_volume_max_ticker As String
         total_stock_volume_max_ticker = " "
        
         
         ' Keep track of the location for each ticker symbol in the summary table

         Dim Summary_Table_Row As Long
         Summary_Table_Row = 2
        
        
         ' Determine Last row

         Dim LastRow As Long
         Dim i As Long
        
         LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

         ' Insert summary table in all worksheets

        If Summary_Table_Row Then

                ' Insert summary table in current worksheet

                ws.Range("J1").Value = "Ticker"
                ws.Range("K1").Value = "Yearly Change"
                ws.Range("L1").Value = "Percent Change"
                ws.Range("M1").Value = "Total Stock Volume"

                ' Insert new summary table for max values

                ws.Range("P2").Value = "Greatest % Increase"
                ws.Range("P3").Value = "Greatest % Decrease"
                ws.Range("P4").Value = "Greatest Total Volume"
                ws.Range("Q1").Value = "Ticker"
                ws.Range("R1").Value = "Value"
        Else
                ' reset summary table for the other sheets

                Summary_Table_Row = True

        End If
        
        ' Set value of initial open stock price
        
        YrOp = ws.Cells(2, 3).Value
        
        ' Loop through all opening and closing stock per ticker symbol
        For i = 2 To Lastrow
        
      
             ' Check if we are still within the same ticker,
             ' if it is not then print into summary table

            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                 ' Set the ticker name

                 Ticker = ws.Cells(i, 1).Value
                    
                 ' Calculate DfInRow and Percent Change

                 YrCl = ws.Cells(i, 6).Value
                 DfInRow = YrCl - YrOp

                 ' condition if YrOp is not zero
                  

                If YrOp <> 0 Then

                     Percent_Change = (DfInRow / YrOp) * 100
                
                End If
                
                 ' Add difference to total stock volume

                 total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
              
                
                 ' Print the Ticker symbol in the summary table

                 ws.Range("J" & Summary_Table_Row).Value = Ticker

                 ' Print the yearly change in the summary table

                 ws.Range("K" & Summary_Table_Row).Value = DfInRow


                 ' Format yearly change cells to reflect positive(green) and negative(red) change

                If (DfInRow > 0) Then

                     ' Green for positive change

                     ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 4

                ElseIf (DfInRow <= 0) Then

                     ' Red for zero value and negative change

                     ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
                End If
                
                 ' Insert percent change into summary table

                 ws.Range("L" & Summary_Table_Row).Value = (CStr(Percent_Change) & "%")

                 ' Insert total stock volume into summary table

                 ws.Range("M" & Summary_Table_Row).Value = total_stock_volume
                
                 ' Add 1 to the summary table row count

                 Summary_Table_Row = Summary_Table_Row + 1

                 ' Reset DfInRow to start with new cells

                 DfInRow = 0

                 ' Reset YrCl to calculate new DfInRow

                 YrCl = 0

                 ' Assign value to next open price

                 YrOp = ws.Cells(i + 1, 3).Value
              
                
                 ' Calculate max values in new summary table for previously defined max variable
                

                If Percent_Change > Percent_Change_max  Then

                     Percent_Change_max = Percent_Change
                     Ticker_max = Ticker

                ElseIf Percent_Change < Percent_Change_min Then

                     Percent_Change_min = Percent_Change
                     Ticker_min = Ticker
                End If
                       
                If total_stock_volume > total_stock_volume_max Then

                     total_stock_volume_max = total_stock_volume
                     total_stock_volume_max_ticker = Ticker
                End If
                
                 ' Reset yearly change and total stock volume

                 DfInRow = 0
                 total_stock_volume = 0
                
            
            ' If the cell in the immediate row has same ticker symbol add to the total stock volume
            
            Else
                 ' Add to total stock volume

                 total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
            End If
            
      
        Next i

             ' Bonus:
             ' First check values not in first spreadsheet
             ' Insert max values in new spread sheet
            
            If Not COMMAND_SPREADSHEET Then
            
                 ws.Range("R2").Value = (CStr(Percent_Change_max) & "%")
                 ws.Range("R3").Value = (CStr(Percent_Change_min) & "%")
                 ws.Range("R4").Value = total_stock_volume_max
                 ws.Range("Q2").Value = Ticker_max
                 ws.Range("Q3").Value = Ticker_min
                 ws.Range("Q4").Value = total_stock_volume_max_ticker
                
            Else
                 COMMAND_SPREADSHEET = False
            End If
        
     Next ws
End Sub