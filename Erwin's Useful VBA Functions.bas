Attribute VB_Name = "Functions"
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' ERWIN'S USEFUL FUNCTIONS
'
' Last update: 2/25/2019
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Option Explicit

' Constants
Public Const vbDQ As String = """"

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Returns the last row of a given column in a given sheet
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Function LastRow(sh As Worksheet, ColumnLet As Variant)
    
    LastRow = sh.Cells(sh.Rows.Count, ColumnLet).End(xlUp).Row
    
End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Returns the last column of a given row in a given sheet
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Function LastCol(sh As Worksheet, RowNum As Variant)

    LastCol = sh.Cells(RowNum, sh.Columns.Count).End(xlToLeft).Column

End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Resets the Excel environment
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Public Sub ResetEnv()
Attribute ResetEnv.VB_ProcData.VB_Invoke_Func = "Q\n14"

    With Application
        .EnableEvents = True
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
    End With
    
    Application.StatusBar = False

End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Check if a file exists given the full path & filename
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Public Function CheckFile(fileName As String) As Boolean

    On Error GoTo errorHandler

    If Dir(fileName) <> "" Then
        CheckFile = True
    Else
        CheckFile = False
    End If
    
exitHandler:

    Exit Function
    
errorHandler:
    
    CheckFile = False

    Resume exitHandler
    
End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Check if a sheet exists given a workbook
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Function SheetExists(shtName As String, Optional wb As Workbook) As Boolean
    
    Dim sht As Worksheet

    If wb Is Nothing Then Set wb = ThisWorkbook
    
    On Error Resume Next
    
    Set sht = wb.Sheets(shtName)
    
    On Error GoTo 0
    
    SheetExists = Not sht Is Nothing

End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Check if a given lat/long coordinate exists within a given polygon
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Public Function PtInPoly(xCoord As Double, yCoord As Double, Polygon As Variant) As Variant

  Dim x As Long, NumSidesCrossed As Long, m As Double, b As Double, Poly As Variant
  
  Poly = Polygon
  
  If Not (Poly(LBound(Poly), 1) = Poly(UBound(Poly), 1) And _
        Poly(LBound(Poly), 2) = Poly(UBound(Poly), 2)) Then
        
    If TypeOf Application.Caller Is Range Then
      PtInPoly = "#UnclosedPolygon!"
    Else
      Err.Raise 998, , "Polygon Does Not Close!"
    End If
    
    Exit Function
    
  ElseIf UBound(Poly, 2) - LBound(Poly, 2) <> 1 Then
  
    If TypeOf Application.Caller Is Range Then
      PtInPoly = "#WrongNumberOfCoordinates!"
    Else
      Err.Raise 999, , "Array Has Wrong Number Of Coordinates!"
    End If
    
    Exit Function
    
  End If
  
  For x = LBound(Poly) To UBound(Poly) - 1
  
    If Poly(x, 1) > xCoord Xor Poly(x + 1, 1) > xCoord Then
      m = (Poly(x + 1, 2) - Poly(x, 2)) / (Poly(x + 1, 1) - Poly(x, 1))
      b = (Poly(x, 2) * Poly(x + 1, 1) - Poly(x, 1) * Poly(x + 1, 2)) / (Poly(x + 1, 1) - Poly(x, 1))
      If m * xCoord + b > yCoord Then NumSidesCrossed = NumSidesCrossed + 1
    End If
    
  Next
  
  PtInPoly = CBool(NumSidesCrossed Mod 2)
  
End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Closes the VBE window
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Sub CloseMainWindow()

Application.VBE.MainWindow.Visible = False

End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Returns the letter equivalent of a column number
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Public Function wColNm(ColNum)

    wColNm = Split(Cells(1, ColNum).Address, "$")(1)
    
End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Checks if value is in array
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Function IsInArray(stringToBeFound As Variant, arr As Variant) As Boolean
  
  IsInArray = (UBound(Filter(arr, stringToBeFound)) > -1)

End Function



