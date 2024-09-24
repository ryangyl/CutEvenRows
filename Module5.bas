Attribute VB_Name = "Module5"
Option Explicit
Sub Alt()
Dim y() As Variant
y = Array("Oct", "Nov")
Dim x As Integer
Dim nc As Integer
Dim b As Integer
Dim a As String
Dim c As Integer

For x = 0 To 1
    Worksheets(y(x)).Activate
    nc = WorksheetFunction.CountA(Range("1:1"))
    a = ActiveSheet.Name
    For b = 1 To nc
    If b Mod 2 = 0 Then
    Cells(1, b).EntireColumn.Cut
    Worksheets(a & " SS").Activate
    Range("A1").Select
    c = b \ 2
    ActiveCell.Offset(0, c).EntireColumn.Select
    ActiveSheet.Paste
    Worksheets(y(x)).Activate
    
    End If
    Next b
    Next x
End Sub
