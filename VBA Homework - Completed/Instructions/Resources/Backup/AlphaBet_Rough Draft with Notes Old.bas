Attribute VB_Name = "Module1"
'Option Explicit


'Variables Set
Sub AlphaBet()
Dim AlphaBets As Workbook
Set AlphaBets = ActiveWorkbook
                                                                                'Dim ws As Worksheet
Dim wsCount As Integer
wsCount = ActiveWorkbook.Worksheets.Count
Dim H As Integer
Dim I As Long
                                                                                'Dim J As Integer

'--------------------'--------------------'--------------------'--------------------'--------------------'--------------------
    
                                                                                
'Start of Worksheet Loop to cycle through all WorkSheets
For H = 1 To wsCount
    Worksheets(H).Activate
                                                                                
                                                                                'For Each ws In Worksheets
                                                                                'WorksheetName = ActiveSheet.Name
                                                                                'MsgBox ("This is " + WorksheetName)
    
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

                                                                                'Autofit to display data
                                                                                'Columns("A:Q").AutoFit
   
                                                                                'Get the Worksheet Name
                                                                                'Dim WorksheetName As String
                                                                                'MsgBox ("This is " + ws.Name)
                                                                                'MsgBox ("This is " + ActiveWorkbook.Worksheets(H).Name)
                                                                                'Call Clear
                                                                                'Next H
 
                                                                                'End Sub
                                                                                'Sub Clear()
                                                                                
                                                                                'Range("H1").Select
                                                                                'Range(Selection, Selection.End(xlDown)).Select
                                                                                'Range(Selection, Selection.End(xlToRight)).Select
                                                                                'Selection.ClearContents
                                                                                'ActiveCell.FormulaR1C1 = " "
                                                                                'Range("H1").Select
                                                                                
                                                                                'Application.Wait (Now + TimeValue("00:00:02"))
                                                                                
                                                                                'Next H
                                                                                'Next ws
                                                                                
                                                                                'Next H
                                                                                
                                                                                'End Sub
                                                                                
                                                                                'Sub Test()
                                                                                
                                                                                
                                                                                'Err.Raise 666
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
    
    
                                                                                    'Go to the Last Row (+1 Cell Down)
                                                                                    'Dim EndofRows As Long
                                                                                    'EndofRows = Cells(Rows.Count, "B").End(xlUp).Select
                                                                                    'ActiveCell.Offset(1, 0).Activate
                                                                                    'ActiveCell.FormulaR1C1 = ""
                                                                                    
                                                                                    'Go to Last Column(+1 Cell Right)
                                                                                    'Dim EndofCols As Long
                                                                                    'EndofCols = Cells("1", Columns.Count).End(xlToLeft).Select
                                                                                    'ActiveCell.Offset(0, 1).Activate
                                                                                    'ActiveCell.FormulaR1C1 = ""
                                                                                    
                                                                                    'Test to make sure LastRow & LastCol are properly counting
                                                                                    'MsgBox ("Last Row: " & LastRow & vbNewLine & "Last Column: " & LastCol)
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
                    If TickerClose <> 0 Then
                        Range("K" & TickerLocation).Value = ((TickerClose - TickerOpen) / TickerOpen)
                                                    
                    Else
                        Range("K" & TickerLocation).Value = "0"
                    
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
    
    
'MISC Code Snippets

'----------------------------------------------------
'----------------------------------------------------
'----------------------------------------------------



'MsgBox ("Fixes Complete")
'Next ws
'MsgBox ("Fixes Complete")

'----------------------------------------------------

'Range("A1").Select
'Selection.End(xlDown).Select
'ActiveCell.Offset(1, 0).Activate
'ActiveCell.FormulaR1C1 = ""

'----------------------------------------------------

'ActiveCell.FormulaR1C1 = "Test1"
'lastRow = combined_sheet.Cells(Rows.Count, "A").End(xlUp).Row + 1
'LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
'LastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column
'Range(Selection, Selection.End(xlDown)).Select
'Range(Selection, Selection.End(xlToRight)).Select
'Range("X#").Select

'----------------------------------------------------


' Range(Selection, Selection.End(xlDown)).NumberFormat = #,###

'Range(Selection, Selection.End(xlDown)).Select
'Selection.NumberFormat = "0.00%" | "#,###"



'----------------------------------------------------


'Worksheet Loop Function
'For Each ws In Worksheets
' CODE
'Next ws
'MsgBox ("Fixes Complete")
'End Sub

'Use to calculate the total
'Total = Range("X#").Value * (1 + Range("Y#").Value in((%))


'----------------------------------------------------

'    Sheets.Add After:=ActiveSheet
'   ActiveSheet.Next.Select
'  Selection.End(xlToLeft).Select
' Selection.End(xlUp).Select
'Range(Selection, Selection.End(xlDown)).Select
'Range(Selection, Selection.End(xlToRight)).Select
'Selection.Copy
'ActiveSheet.Previous.Select
' ActiveSheet.Paste
'  Range("A1").Select
'   Selection.End(xlDown).Select
'    Range("A33").Select

'----------------------------------------------------

'         ' Add a sheet named "Combined Data"
'   Sheets.Add.Name = "Combined_Data"
'           ' Move created sheet to be first sheet
'     Sheets("Combined_Data").Move Before:=Sheets(1)
'             ' Specify the location of the combined sheet
'       Set combined_sheet = Worksheets("Combined_Data")
'
'             ' Loop through all sheets
'    For Each ws In Worksheets
'
'         ' Find the last row of the combined sheet after each paste
'          ' Add 1 to get first empty row
'       lastRow = combined_sheet.Cells(Rows.Count, "A").End(xlUp).Row + 1
'
'         ' Find the last row of each worksheet
'          ' Subtract one to return the number of rows without header
'       lastRowState = ws.Cells(Rows.Count, "A").End(xlUp).Row - 1
'
'          ' Copy the contents of each state sheet into the combined sheet
'       combined_sheet.Range("A" & lastRow & ":G" & ((lastRowState - 1) + lastRow)).Value = ws.Range("A2:G" & (lastRowState + 1)).Value
'
'    Next ws
'
'       ' Copy the headers from sheet 1
'     combined_sheet.Range("A1:G1").Value = Sheets(2).Range("A1:G1").Value
'
'          ' Autofit to display data
'        combined_sheet.Columns("A:G").AutoFit



' ----------------------------------------------------


'        'Dim newPrice As Double
'       newPrice = budget / (1 + Range("H3").Value)
'      MsgBox ("The New Price is " + Str(newPrice) + " !")
'   'Use a worksheet function to round the new price down to the nearest whole dollar
'   newPrice = Application.WorksheetFunction.RoundDown(newPrice, 0)
'
'     ' Change the price
'      Range("F3").Value = newPrice
'       ' Change the new total
'        Range("L3").Value = newPrice * (1 + Range("H3").Value)

'------------------------------------------------
' Dim Sheetname As String
'(Sheetname) = Split(ActiveSheet.Name,",")(0)
' Msgbox Sheetname

'------------------------------------------------

''=RIGHT(N2,LEN(N2)-SEARCH("/",N2,1))
'split(ActiveSheet.Name,”_”)(0)
' "\Wells_Fargo"

'------------------------------------------------

' Determine the Last Row & Column
'LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
'LastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column

' Grabbed the WorksheetName
'WorksheetName = ws.Name

' Split the WorksheetName
'SplitWord = Split(WorksheetName, "_")

' Add/Insert a Column at location
'ws.Range("A1").EntireColumn.Insert

' Add word to the First Column Header
'ws.Cells(1, 1).Value = "  "

' Add (X part/0) of 'SplitWord' to all rows in "X" column til end row
'ws.Range("A2:A" & LastRow) = SplitWord(0)

'----------------------------------------------------

' Determine the Last Column Number
'LastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column
' Rename the XX columns by looping through and renaming each
'        For i = 3 To LastColumn
'           YearHeader = ws.Cells(1, i).Value
'          YearSplit = Split(YearHeader, " ")
'      This is too extract a search/split value/word with delimiter (" ")
'           ws.Cells(1, i).Value = YearSplit(3)
'              The(x) function choose the array to return starting from 0,1,2,...
'       Next i

'----------------------------------------------------


' Add/Format a Cell with loop in whole Worksheet
'   For i = 2/(x) To LastRow
'       For j = 2/(x) To LastColumn
'          ws.Cells(i, j).Style = "Currency","Number","Date","Time"....etc
'       Next j
'   Next i


'------------------------------------------------

' Set the Font color to Red
'  Range("XX").Font.ColorIndex = 3
'   Range("XX").Font.Color = vbRed
'    Range("XX").Font.Color = RGB(255, 0, 0)
' Set the Cell Colors to Red (3 Different Options)
' Range("X#:X#").Interior.ColorIndex = 3
'  Range("X#:X#").Interior.Color = vbRed
'   Range("X#:X#").Interior.Color = RGB(255, 0, 0)
'             ---------------
' Set the Font Color to Green
' Range("X#").Font.ColorIndex = 4
'  Range("X#").Font.Color = vbGreen
'   Range("X#").Font.Color = RGB(0, 255, 0)
' Set the Cell Colors to Green
'    Range("X#:X#").Interior.ColorIndex = 4
'   Range("X#:X#").Interior.Color = vbGreen
'  Range("X#:X#").Interior.Color = RGB(0, 255, 0)

'------------------------------------------------

'Option Explicit
'Dim Global_Variable As Integer
'This variable is a global variable and can be used by any Sub routine

'Sub xxx()
'Dim Local_Variable As Integer
'This variable is a Local variable and can only be used by this Sub routine

' Check if the value is greater than or equal to x...
'If Range("B2").value >= XX Then
'If Cells(x, x).Value >= XX Then

' Establish that X is True
'Range("XX").value = "XXXX"
'Cells(X, X).Value = "XXXX"

'Color the X cell green
'Range("XX").Interior.Color = vbGreen
'Cells(X, X).Interior.ColorIndex = 4

'Color the X cell red
'Range("XX").Interior.Color = vbRed
'Cells(X, X).Interior.ColorIndex = 3

'Set X to "XX"
'Cells(X, X).Value = "X"

'------------------------------------------------

'For r = 1 To 8
'For c = 1 To 8
'  If (r + c) Mod 2 = 0 Then
'       Cells(r, c).Interior.ColorIndex = 1 ' Black
'    Else
'         Cells(r, c).Interior.ColorIndex = 3 ' Red
'      End If
'   Next c
'Next r



'------------------------------------------------
'See SentenceBreaker()
'See MacroTest_EndRow() / Star Counter with VBA
'See Splitting
'See formatterITH
'See Credit card
'See Wells Fargo 1&2
'See Internet



'Written by: Ithamar Francois
