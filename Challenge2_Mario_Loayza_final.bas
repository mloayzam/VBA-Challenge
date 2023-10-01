Attribute VB_Name = "Module1"
Sub year()

'Challenge 2 VBA Scripts - Student: Mario Loayza

' ----------------------------------------------------------------
' Inside each Year worksheet output: Ticker Symbol, Yearly price,
' percent change and total stock volumen calculation
' Set the Percentge formating for percentage change
' Conditional formating for positive change and negative change
' Calculation: "Greatest % increase", "Greatest % decrease", and "Greatest total volume"
' ----------------------------------------------------------------


Dim Year_sheet As Worksheet

For YR = 2018 To 2020

Set Year_sheet = Worksheets(CStr(YR))
'Set Year_sheet = Worksheets("2018")

Year_sheet.Cells(1, 9).Value = "Ticker"
Year_sheet.Cells(1, 10).Value = "Yearly Change"
Year_sheet.Cells(1, 11).Value = "Percent Change"
Year_sheet.Cells(1, 12).Value = "Total Stock Volume"

Year_sheet.Range("I1:L1").Font.Bold = True

  ' Set an initial variable for holding the ticker name and calculation variables
  Dim ticker_name As String
  Dim L1 As Double
  Dim L2 As Double
  Dim opy As String
  Dim cly As String
  Dim ndata As Long
  
  'openyear, close year and number of data
   opy = CStr(YR) & "0102"
   cly = CStr(YR) & "1231"
  
  ndata = 800000
  


  ' Set an initial variable for holding the total Stock volume
   Dim Volume_Total As Double
   Volume_Total = 0
        
  ' Keep track of the location for each ticker name in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2



  ' Loop through all tickers
  For i = 2 To ndata

    ' Check if we are still within the same ticker name, if it is not...
    If Year_sheet.Cells(i + 1, 1).Value <> Year_sheet.Cells(i, 1).Value Then

      ' Set the ticker name
      ticker_name = Year_sheet.Cells(i, 1).Value

      ' Add to the Stock Volume Total
      Volume_Total = Volume_Total + Year_sheet.Cells(i, 7).Value

      ' Print the ticker name in the Summary Table
      Year_sheet.Range("I" & Summary_Table_Row).Value = ticker_name

      ' Print the Stock Volume Amount to the Summary Table
      Year_sheet.Range("L" & Summary_Table_Row).Value = Volume_Total

      
      'Condition for looking at the cell for the closing price given the end of year date of the ticker name
      If Year_sheet.Cells(i, 2).Value = cly Then
    
      L2 = Year_sheet.Cells(i, 6).Value
   

      End If
    
      'Print the claculation of yearly change and percent change
      Year_sheet.Range("J" & Summary_Table_Row).Value = L2 - L1

      Year_sheet.Range("K" & Summary_Table_Row).Value = (L2 - L1) / L1

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Stock Volume Total
      Volume_Total = 0
  


   
    ' If the cell immediately following a row is the same ticker...
    Else

      ' Add to the Stock Volume Total
      Volume_Total = Volume_Total + Year_sheet.Cells(i, 7).Value

    End If


    'Condition for looking at the cell for the opening price given the beginning of year date of the ticker name

    If Year_sheet.Cells(i, 2).Value = opy Then
            
         L1 = Year_sheet.Cells(i, 3).Value

    End If
    

 
Next i
    
        ' --------------------------------------------
        ' INCLUDE THE PERCENT FORMAT
        ' --------------------------------------------

        ' Add percent for cells
        For i = 2 To ndata

            ' For columns Percent Change only
            For j = 11 To 11

                Year_sheet.Cells(i, j).Style = "Percent"

            Next j

        Next i
    
    
         ' --------------------------------------------------------------------------
         ' CONDITIONAL FORMATING FOR HIGHLIGHTING POSITIVE CHANGE AND NEGATIVE CHANGE
         ' ---------------------------------------------------------------------------
        
        'Conditional formating
        For i = 2 To ndata
        
            If Year_sheet.Cells(i, 11).Value > 0 Then
        
            ' Format color to green
            Year_sheet.Cells(i, 11).Interior.ColorIndex = 4
        
             Else
        
            'Formar color to red
            Year_sheet.Cells(i, 11).Interior.ColorIndex = 3
        
            End If
        
        Next i
        
        
         ' ---------------------------------------------------------------------------------------
         ' Calculation: "Greatest % increase", "Greatest % decrease", and "Greatest total volume"
         ' ----------------------------------------------------------------------------------------
         
    Year_sheet.Cells(2, 14).Value = "Greatest%Increase"
    Year_sheet.Cells(3, 14).Value = "Greatest%decrease"
    Year_sheet.Cells(4, 14).Value = "GreatestTotalVolume"

    Year_sheet.Cells(1, 15).Value = "Ticker"
    Year_sheet.Cells(2, 16).Style = "Percent"
    Year_sheet.Cells(1, 16).Value = "Value"
    Year_sheet.Cells(3, 16).Style = "Percent"


    Year_sheet.Cells(2, 16).Value = WorksheetFunction.Max(Year_sheet.Range("K2:K800000"))
    Year_sheet.Cells(3, 16).Value = WorksheetFunction.Min(Year_sheet.Range("K2:K800000"))
    Year_sheet.Cells(4, 16).Value = WorksheetFunction.Max(Year_sheet.Range("L2:L800000"))
        

'Print the ticker name according with the calculations

    Dim Max  As Double
    Dim ticker As String
    Dim Min  As Double
    Dim GT   As Double


        For i = 2 To 5000
        ticker = Year_sheet.Cells(i, 9).Value
        Max = Year_sheet.Cells(i, 11).Value
        Min = Year_sheet.Cells(i, 11).Value
        GT = Year_sheet.Cells(i, 12).Value

        If Max = Year_sheet.Cells(2, 16).Value Then

            Year_sheet.Cells(2, 15).Value = ticker

        End If

        If Min = Year_sheet.Cells(3, 16).Value Then

            Year_sheet.Cells(3, 15).Value = ticker

        End If

        If GT = Year_sheet.Cells(4, 16).Value Then

            Year_sheet.Cells(4, 15).Value = ticker

        End If

        Next i
        
Next YR
        
End Sub
