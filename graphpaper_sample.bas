Attribute VB_Name = "graphpaper_sample"
'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'1pt=1/72inch 1inch=25.4mm
'1pt=25.4mm/72inch a.0.3528mm
'1mm=72inch/25.4mm a.2.8346pt
'  cells_graphpaper: 1point 0.3528mm 'ratio = 0.0685
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub cells_graphpaper()
Dim sname As String: sname = "graphpaper"
Dim rc As Variant: rc = Array(2, 2, 5, 10)
Dim cm As Variant: cm = 5 'mm unit
Dim a As Double, ratio As Double
Dim pt As Double, mm As Double, inch As Double
pt = 1: mm = 25.4: inch = 72: ratio = 0.0001
a = (inch / mm) * cm
Call addsheet(sname)
Application.ScreenUpdating = False
With ThisWorkbook
  With .Worksheets(1)
    'cm = Application.CentimetersToPoints(1)
    .Cells.Clear
    .Cells.RowHeight = a
    Do While ActiveCell.Width <> ActiveCell.Height
      ratio = ratio + 0.0001
      .Cells.ColumnWidth = a * ratio
    Loop
    Debug.Print "ƒZƒ‹• " & cm & "mm ŠÔŠu‚Å‚·"
    a = Round(a)
    With .Range(.Cells(rc(0), rc(1)), .Cells(rc(2), rc(3)))
      .Interior.Color = RGB(200, 240, 250)
      .Borders.LineStyle = xlContinuous
    End With
  End With
End With
Application.ScreenUpdating = True
End Sub
'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'  addsheet:
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub addsheet(Optional sname As String = "")
Call delsheet(CVar(sname))
With ThisWorkbook
  .Worksheets.Add.Name = sname
End With
End Sub
'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'  delsheet:
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub delsheet(Optional sname As Variant)
Dim s As Variant, i As Integer
sname = Split(sname, ",")
Application.DisplayAlerts = False
With ThisWorkbook
  For i = 0 To UBound(sname)
    For Each s In .Worksheets
      If s.Name = sname(i) Then s.Delete
    Next
  Next
End With
Application.DisplayAlerts = True
End Sub
