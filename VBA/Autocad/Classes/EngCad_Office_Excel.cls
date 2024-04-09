VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EngCad_Office_Excel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Function EliminarFilasEnBlanco(Range1 As String)
Dim r As Range, rows As Long, i As Long
Set r = ActiveSheet.Range(Range1)
rows = r.rows.Count
For i = rows To 1 Step (-1)
If WorksheetFunction.CountA(r.rows(i)) = 0 Then r.rows(i).Delete
Next
Range("A1").Select
End Function
Public Sub OcultarAplicacion()
Application.Visible = False
End Sub
Public Sub MostrarAplicacion()
Application.Visible = True
End Sub
Public Sub DeleteAllShapes()
Dim Shp As Shape
For Each Shp In ActiveSheet.Shapes
Shp.Delete
Next Shp
End Sub
Public Function EliminarHojasEnBlanco()
Application.DisplayAlerts = False
For Each Worksheet In Worksheets
If IsEmpty(Worksheet.UsedRange) Then
Worksheet.Delete
End If
Next
Application.DisplayAlerts = True
End Function
Public Function NuevaHoja(nombre As String)
Application.DisplayAlerts = False
For Each Worksheet In Worksheets
If Worksheet.Name = nombre Then
Worksheet.Delete
Else:
End If
Next
Sheets.Add After:=ActiveSheet
ActiveSheet.Name = nombre
Application.DisplayAlerts = True
End Function
Public Function ReemplazarTexto(TextoEntrada As String, TextoSalida As String)
Cells.Replace What:=TextoEntrada, Replacement:=TextoSalida, LookAt:=xlPart, _
SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
ReplaceFormat:=False
End Function
Public Function SeleccionarColumnas(Columna1 As String, Columna2 As String)
Columns(Columna1 & ":" & Columna2).Select
End Function

Public Function ContarCeldasLlenas(Letra As String, Hasta As Integer)
n = Application.WorksheetFunction.CountA(Range(Letra & "1:" & Letra & Hasta))
ContarCeldasLlenas = n
End Function









