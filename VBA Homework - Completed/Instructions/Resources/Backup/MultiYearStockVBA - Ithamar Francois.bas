Attribute VB_Name = "Module1"
'Option Explicit
'Ithamar Francois
'VBA Challenge Homework
'RU Data Bootcamp 2021


'Variables Set
Sub AlphaBet()
Dim AlphaBets As Workbook
Set AlphaBets = ActiveWorkbook
                                                                                
Dim wsCount As Integer
wsCount = ActiveWorkbook.Worksheets.Count
Dim H As Integer
Dim I As Long
                                                                                

'--------------------'--------------------'--------------------'--------------------'--------------------
    
                                                                                
'Start of Worksheet Loop to cycle through all WorkSheets
For H = 1 To wsCount
    Worksheets(H).Activate
                                                                                
                                                                                
    
    '--------------------
    'Creating Headers for Columns for each Worksheet
     Range("I1").Value = " Ticker "
     Range("J1").Value = " Yearly Change ($) "
     Range("K1").Value = " Percent Change (%)"
     Range("L1").Value = " Total Stock Volume "
     Range("M1").Value = " Yearly Opening Price "
     Range("N1").Value = " Yearly Closing Price "
     Range("O2").Value = "   Greatest % Increase "
     Range("O3").Value = "   Greatest % Decrease "
     Range("O4").Value = "   Greatest Total Volume "
     Range("P1").Value = " Ticker "
     Range("Q1").Value = " Value "

                                                                               
    '--------------------
   
    'Set variable for Ticker Symbol
     Dim Ticker As String
      Ticker = " "
    
    'Set variable for holding Total Ticker Volume
      Dim TickerVolume As Double
        TickerVolume = 0

    'Keep track of the location for each Ticker
      Dim TickerLocation As Double
        TickerLocation = 2
  
    'Keep track of the time this year a Ticker has been trading
      Dim TickerYear As Integer
        TickerYear = 1
    
    'Set variable for holding Ticker new year's opening price
      Dim TickerOpen As Double
       TickerOpen = "0.00"
    
    'Set variable for holding Ticker new year's closing price
      Dim TickerClose As Double
       TickerClose = "0.00"
    
  '--------------------
   
   'Set variables for holding BONUS values
     Dim GreatestIncrease As Double
      GreatestIncrease = 0
     
     Dim GreatestDecrease As Double
      GreatestDecrease = 0
     
     Dim GreatestVolume As Double
      GreatestVolume = 0
     
     Dim TickerBonusIncrease As String
      TickerBonusIncrease = " "
     
     Dim TickerBonusDecrease As String
      TickerBonusDecrease = " "
     
     Dim TickerBonusVolume As String
      TickerBonusVolume = " "
     
   '--------------------
   
    'Determine the Last Row & Last Column length
     Dim LastRow As Long
      LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
     Dim LastCol As Long
      LastCol = Cells(1, Columns.Count).End(xlToLeft).Column
    
    
                                                                                   
        '--------------------
        'Formatting of cells/columns that will hold different values types
        
        Range("C1").Select
        Range(Selection, Selection.End(xlDown)).NumberFormat = "$#,##0.00"
        
        Range("D1").Select
        Range(Selection, Selection.End(xlDown)).NumberFormat = "$#,##0.00"
        
        Range("E1").Select
        Range(Selection, Selection.End(xlDown)).NumberFormat = "$#,##0.00"
        
        Range("F1").Select
        Range(Selection, Selection.End(xlDown)).NumberFormat = "$#,##0.00"
        
        
        Range("J1").Select
        Range(Selection, Selection.End(xlDown)).NumberFormat = "$#,##0.00"
        
        Range("K1").Select
        Range(Selection, Selection.End(xlDown)).NumberFormat = "0.00%"
        
        
        Range("L1").Select
        Range(Selection, Selection.End(xlDown)).NumberFormat = "0.00E+00"
        
        
        Range("M1").Select
        Range(Selection, Selection.End(xlDown)).NumberFormat = "$#,##0.00"
        
        Range("N1").Select
        Range(Selection, Selection.End(xlDown)).NumberFormat = "$#,##0.00"
             
'--------------------'-------------------- '-------------------- '-------------------- '--------------------

'Start of Loop and Code Execution
Range("A2").Select
For I = 2 To LastRow
'Loop through all stocks/Tickers to check if we are still comparing the same Tickers/Year
    
    If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
    'If a Ticker (immediately following previous row's Ticker) is not the same, Then...
    
            'Copy Ticker name
             Ticker = Cells(I, 1).Value

            'Add the day's Volume to Ticker Total Volume
             TickerVolume = TickerVolume + Cells(I, 7).Value

            'Print the Ticker Symbol in the Ticker Location Column
             Range("I" & TickerLocation).Value = Ticker

            'Print the Total Ticker Volume to the Total Stock Volume Column
             Range("L" & TickerLocation).Value = TickerVolume

            'Reset the Ticker Volume Total
             TickerVolume = 0
             
            'Returns Ticker Opening Price for Year
             TickerOpen = Cells(I - (TickerYear - 1), 3).Value
             Range("M" & TickerLocation).Value = TickerOpen
             'Cells((i - (TickerYear - 1)), 3).Select
             TickerYear = 1
             
            'Returns Ticker Closing Price for Year
             TickerClose = Cells(I, 6).Value
             Range("N" & TickerLocation).Value = TickerClose
             
            'Calculate the Yearly Change in Stock Price for Year
             Range("J" & TickerLocation).Value = (TickerClose - TickerOpen)
                    
                    'Color code cell to show if values of Yearly & Percent Change are postive/negative
                    If Range("J" & TickerLocation).Value >= 0 Then
                            Range("J" & TickerLocation).Interior.Color = vbGreen
                            Range("K" & TickerLocation).Interior.Color = vbGreen
                            
                    ElseIf Range("J" & TickerLocation).Value < 0 Then
                            Range("J" & TickerLocation).Interior.Color = vbRed
                            Range("K" & TickerLocation).Interior.Color = vbRed
                    
                    End If
                     
                             
            'Calculate the Percent Change in opening to closing Stock Price for Year
                    'If ClosingPrice is 0, then use if statement to return "0" (to avoid divide by 0 error)
                    If TickerOpen = 0 Then
                        Range("K" & TickerLocation).Value = "0"
                        
                    Else
                        Range("K" & TickerLocation).Value = ((TickerClose - TickerOpen) / TickerOpen)
                                                    
                    End If
            
            '--------------------
            
            'BONUS Loops comparing Values in "Column K" against values above it (like 'For Loop' above with Ticker)
            If (Range("K" & TickerLocation).Value > Range("K" & (TickerLocation - 1)).Value And _
                (Range("K" & TickerLocation).Value > GreatestIncrease)) Then
                
                GreatestIncrease = Range("K" & TickerLocation).Value
                TickerBonusIncrease = Range("I" & TickerLocation).Value
                          
            End If
                
                
                
            If (Range("K" & TickerLocation).Value < Range("K" & (TickerLocation - 1)).Value And _
                (Range("K" & TickerLocation).Value < GreatestDecrease)) Then
                    
                GreatestDecrease = Range("K" & TickerLocation).Value
                TickerBonusDecrease = Range("I" & TickerLocation).Value
            
            End If
                
                
                
                'Bonus Loop comparing Values in "Column L" against values above it (like 'For Loop' with Ticker)
                If (Range("L" & TickerLocation).Value > Range("L" & (TickerLocation - 1)).Value And _
                (Range("L" & TickerLocation).Value > GreatestVolume)) Then
                    
                    GreatestVolume = Range("L" & TickerLocation).Value
                    TickerBonusVolume = Range("I" & TickerLocation).Value
                
                End If
                
            '--------------------
                              
                        
             
           'Add one to Ticker Location tracker to prep next symbol's entry
             TickerLocation = TickerLocation + 1
    
    
        '--------------------
             
    Else
    'If a Ticker immediately following a previous entry is the same symbol...
    
            'Add to the Ticker Volume Total
             TickerVolume = TickerVolume + Cells(I, 7).Value
             
            'Add to Ticker Year tracker
             TickerYear = TickerYear + 1

    End If
    '--------------------

'Restart loop's For statement again until "i" is exhausted at 'LastRow'
Next I
            
            
'-------------------- '-------------------- '-------------------- '-------------------- '--------------------
            
            
            'Format BONUS table for results
             Range("Q2:Q3").Select
             Range(Selection, Selection.End(xlDown)).NumberFormat = "0.00%"
            
             Range("Q4").NumberFormat = "0.0000E+00"
            
                        
             'Enter BONUS data in table
              Range("P2").Value = TickerBonusIncrease
              Range("Q2").Value = GreatestIncrease
            
              Range("P3").Value = TickerBonusDecrease
              Range("Q3").Value = GreatestDecrease
            
              Range("P4").Value = TickerBonusVolume
              Range("Q4").Value = GreatestVolume
            
            
       '--------------------
        'Data Sheets formatted for view and End
        
         Range("H1").Select
            
        'Autofit to display data
         Columns("A:Q").AutoFit
         Range("O1:Q4").Select
 
Next H
'End of Worksheet Loop

'Resets cursor positions after macro for viewing
Worksheets(1).Activate
Range("H1").Select

Range("O1:Q4").Select
MsgBox ("Data is ready for view")


'Finished!
End Sub
    
'----------------------------------------------------'----------------------------------------------------

'Written by: Ithamar Francois
