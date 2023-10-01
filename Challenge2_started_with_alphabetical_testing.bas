Attribute VB_Name = "Module1"
Sub Challenge2test()

' -------------------------------------------------
' Combine the data for all sheets in only worksheet
' -------------------------------------------------

'from activity 3, 08-Stu_Census_Pt2

    'combine data

    Sheets.Add.Name = "combined"

    'move created sheet to be first sheet
    Sheets("combined").Move Before:=Sheets(1)
    
    'location of the combined data

    Set Combined_Sheet = Worksheets("Combined")

    'Loop though all sheets
    For Each ws In Worksheets
    
        'Find the last row of the combined sheet after each paste
        'Add 1 to get first empty row
        LastRow = Combined_Sheet.Cells(Rows.Count, "A").End(xlUp).Row + 1
        'Find the last row of each worksheet
        'Subtract one to return the number of rows without header
        lastRowYear = ws.Cells(Rows.Count, "A").End(xlUp).Row - 1
        'Copy the contents of each year sheet into the combined sheet
        Combined_Sheet.Range("A" & LastRow & ":G" & ((lastRowYear - 1) + LastRow)).Value = ws.Range("A2:G" & (lastRowYear + 1)).Value

    Next ws
    
    'Copy the headers from sheet 1
    Combined_Sheet.Range("A1:G1").Value = Sheets(2).Range("A1:G1").Value

    ' Autofit to display data
    Combined_Sheet.Columns("A:G").AutoFit



' ----------------------------------------------------------------
' Inside combined worksheet output: Ticker Symbol, Yearly price,
' percent change and total stock volumen calculation
' Set the Percentge formating for percentage change
' Conditional formating for positive change and negative change
' Calculation: "Greatest % increase", "Greatest % decrease", and "Greatest total volume"
' ----------------------------------------------------------------

'From activity 3 06-Stu_CreditCardChecker-CellComparison

 Set Combined_Sheet = Worksheets("Combined")

Combined_Sheet.Cells(1, 9).Value = "Ticker"
Combined_Sheet.Cells(1, 10).Value = "Yearly Change"
Combined_Sheet.Cells(1, 11).Value = "Percent Change"
Combined_Sheet.Cells(1, 12).Value = "Total Stock Volume"

Worksheets("combined").Range("I1:L1").Font.Bold = True

  ' Set an initial variable for holding the ticker name and calculation variables
  Dim ticker_name As String
  Dim L1 As Double
  Dim L2 As Double
  


  ' Set an initial variable for holding the total Stock volume
   Dim Volume_Total As Double
   Volume_Total = 0
        
  ' Keep track of the location for each ticker name in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2



  ' Loop through all tickers
  For i = 2 To 158885

    ' Check if we are still within the same ticker name, if it is not...
    If Combined_Sheet.Cells(i + 1, 1).Value <> Combined_Sheet.Cells(i, 1).Value Then

      ' Set the ticker name
      ticker_name = Combined_Sheet.Cells(i, 1).Value

      ' Add to the Stock Volume Total
      Volume_Total = Volume_Total + Combined_Sheet.Cells(i, 7).Value

      ' Print the ticker name in the Summary Table
      Combined_Sheet.Range("I" & Summary_Table_Row).Value = ticker_name

      ' Print the Stock Volume Amount to the Summary Table
      Combined_Sheet.Range("L" & Summary_Table_Row).Value = Volume_Total

      
      'Condition for looking at the cell for the closing price given the end of year date of the ticker name
      If Combined_Sheet.Cells(i, 2).Value = 20201231 Then
    
      L2 = Combined_Sheet.Cells(i, 6).Value
   

      End If
    
      'Print the claculation of yearly change and percent change
      Combined_Sheet.Range("J" & Summary_Table_Row).Value = L2 - L1

      Combined_Sheet.Range("K" & Summary_Table_Row).Value = (L2 - L1) / L1

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Stock Volume Total
      Volume_Total = 0
  


   
    ' If the cell immediately following a row is the same ticker...
    Else

      ' Add to the Stock Volume Total
      Volume_Total = Volume_Total + Combined_Sheet.Cells(i, 7).Value

    End If


    'Condition for looking at the cell for the opening price given the beginning of year date of the ticker name

    If Combined_Sheet.Cells(i, 2).Value = 20200102 Then
            
         L1 = Combined_Sheet.Cells(i, 3).Value

    End If
    

 
Next i
    
        ' --------------------------------------------
        ' INCLUDE THE PERCENT FORMAT
        ' --------------------------------------------

        ' Add percent for cells
        For i = 2 To 158885

            ' For columns Percent Change only
            For j = 11 To 11

                Combined_Sheet.Cells(i, j).Style = "Percent"

            Next j

        Next i
    
    
         ' --------------------------------------------------------------------------
         ' CONDITIONAL FORMATING FOR HIGHLIGHTING POSITIVE CHANGE AND NEGATIVE CHANGE
         ' ---------------------------------------------------------------------------
        
        'Conditional formating
        For i = 2 To 158885
        
            If Cells(i, 11).Value > 0 Then
        
            ' Format color to green
            Cells(i, 11).Interior.ColorIndex = 4
        
             Else
        
            'Formar color to red
            Cells(i, 11).Interior.ColorIndex = 3
        
            End If
        
        Next i
        
        
         ' ---------------------------------------------------------------------------------------
         ' Calculation: "Greatest % increase", "Greatest % decrease", and "Greatest total volume"
         ' ----------------------------------------------------------------------------------------
         
Combined_Sheet.Cells(2, 14).Value = "Greatest%Increase"
Combined_Sheet.Cells(3, 14).Value = "Greatest%decrease"
Combined_Sheet.Cells(4, 14).Value = "GreatestTotalVolume"

Combined_Sheet.Cells(1, 15).Value = "Ticker"
Combined_Sheet.Cells(2, 16).Style = "Percent"
Combined_Sheet.Cells(1, 16).Value = "Value"
Combined_Sheet.Cells(3, 16).Style = "Percent"


Combined_Sheet.Cells(2, 16).Value = WorksheetFunction.Max(Range("K2:K158885"))
Combined_Sheet.Cells(3, 16).Value = WorksheetFunction.Min(Range("K2:K158885"))
Combined_Sheet.Cells(4, 16).Value = WorksheetFunction.Max(Range("L2:L158885"))
        

'Print the ticker name according with the calculations

Dim Max  As Double
Dim ticker As String
Dim Min  As Double
Dim GT   As Double


For i = 2 To 700
ticker = Combined_Sheet.Cells(i, 9).Value
Max = Combined_Sheet.Cells(i, 11).Value
Min = Combined_Sheet.Cells(i, 11).Value
GT = Combined_Sheet.Cells(i, 12).Value

If Max = Combined_Sheet.Cells(2, 16).Value Then

Combined_Sheet.Cells(2, 15).Value = ticker

End If

If Min = Combined_Sheet.Cells(3, 16).Value Then

Combined_Sheet.Cells(3, 15).Value = ticker

End If

If GT = Combined_Sheet.Cells(4, 16).Value Then

Combined_Sheet.Cells(4, 15).Value = ticker

End If

Next i
        
        
End Sub





    


