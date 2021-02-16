Attribute VB_Name = "Module1"
Option Explicit
Sub formatter()

' Set the Font color to Red
    Range("A1").Font.ColorIndex = 3
    Range("A1").Font.Color = vbRed
    Range("A1").Font.Color = RGB(255, 0, 0)

' Set the Cell Colors to Red (3 Different Options)
    Range("A2:A5").Interior.ColorIndex = 3
    Range("A2:A5").Interior.Color = vbRed
    Range("A2:A5").Interior.Color = RGB(255, 0, 0)
    
'-------------------------------------------------------------

' Set the Font Color to Green
    Range("B1").Font.ColorIndex = 4
    Range("B1").Font.Color = vbGreen
    Range("B1").Font.Color = RGB(0, 255, 0)
    
' Set the Cell Colors to Green
    Range("B2:B5").Interior.ColorIndex = 4
    Range("B2:B5").Interior.Color = vbGreen
    Range("B2:B5").Interior.Color = RGB(0, 255, 0)

'-------------------------------------------------------------

' Set the Color Index to Blue
    Range("C1").Font.ColorIndex = 5
    Range("C1").Font.Color = vbGreen
    Range("C1").Font.Color = RGB(0, 0, 255)

' Set the Cell Colors to Blue
    Range("C2:C5").Interior.ColorIndex = 5
    
'-------------------------------------------------------------

' Set the Color Index to Magenta
    Range("D1").Font.ColorIndex = 7
' Set the Cell Colors to Magenta
    Range("D2:D5").Interior.ColorIndex = 7
    
    
'Black
Range("E1").Font.ColorIndex = 0
Range("E1:E5").Interior.Color = RGB(0, 0, 0)
Range("E2", "E4").Font.Color = vbEmpty

'White
Range("E1").Font.ColorIndex = 2
Range("E5").Font.Color = RGB(255, 255, 255)
Range("E2:E4").Interior.Color = xlNone

'See www.w3schools.com for color palette/codes
'See this website for color guides: http://dmcritchie.mvps.org/excel/colors.htm
End Sub
