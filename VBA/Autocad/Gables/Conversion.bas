Attribute VB_Name = "Conversion"
Dim Entidad As AcadEntity
Dim i As Integer
Dim modelspace As AcadModelSpace
Dim PaperSpace As AcadPaperSpace
Dim ssetObj As AcadSelectionSet
Dim layerObj As AcadLayer
Dim LayerCol As AcadLayers
Dim Linea As AcadLine
Dim Poly As AcadLWPolyline




Sub Completo()

ConvSplineToArc
End Sub

Sub SplineToPoly2()

ConvierteAPolilinea

End Sub
Sub POlilineaToArco()

PolilineaAArco

End Sub

Sub PintaRojo()

LargoLineas

End Sub

Sub ConvierteAPolilinea()
    Set ssetObj = AutoCAD.Application.ActiveDocument.SelectionSets.Add(Rnd)
    ssetObj.SelectOnScreen
    AutoCAD.Application.ActiveDocument.SendCommand "SPLINEDIT" & vbCr & "p" & vbCr & "p" & vbCr & 10 & vbCr
    ssetObj.Delete
End Sub

Sub LargoLineas()
    Dim first As Variant
    Dim second As Variant
    Dim i As Long
    Range("A2:Z20000").Value = ""
    'activarAcad
    Set modelspace = AutoCAD.Application.ActiveDocument.modelspace
    Dim Linea As AcadLine
    i = 2
    first = AutoCAD.Application.ActiveDocument.Utility.GetPoint(, "Select first corner: ")
    second = AutoCAD.Application.ActiveDocument.Utility.GetCorner(first, "Select second corner: ")
    Set ssetObj = AutoCAD.Application.ActiveDocument.SelectionSets.Add(Rnd)
    ssetObj.Select acSelectionSetWindow, first, second
    On Error Resume Next
    For Each Linea In ssetObj
    Cells(i, 4) = Linea.Length
    Cells(i, 5) = Linea.ObjectID
    i = i + 1
    Next
    ActiveWorkbook.Worksheets("AcDbLine-MS").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("AcDbLine-MS").Sort.SortFields.Add Key:=Range("D2") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("AcDbLine-MS").Sort
        .SetRange Range(Cells(2, 4), Cells(i - 1, 5))
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Dim factor As Double
    factor = AutoCAD.Application.ActiveDocument.Utility.GetReal("Ingrese Factor: ")
    For a = 1 To factor
        For Each Linea In ssetObj
        If Linea.ObjectID = Cells(a + 1, 5) Then Linea.Color = acRed
        Next
    Next
    ssetObj.Delete
    Range("A2:Z20000").Value = ""
End Sub

Sub PolilineaAArco()

    Dim Poly As AcadLWPolyline
    Dim n As Integer
    Set ssetObj = AutoCAD.Application.ActiveDocument.SelectionSets.Add(Rnd)
    ssetObj.SelectOnScreen
    i = 2
        On Error Resume Next
        For Each Poly In ssetObj
        Bound = UBound(Poly.Coordinates)
        x = 0
        Y = 1
        For d = 0 To Bound / 2
        Cells(i, 4) = Round(Poly.Coordinates(x), 4)
        Cells(i, 4) = Replace(Cells(i, 4), ",", ".")
        Cells(i, 5) = Round(Poly.Coordinates(Y), 4)
        If Err Then Err.Clear
        x = x + 2
        Y = Y + 2
        i = i + 1
        Next
        n = Application.WorksheetFunction.CountA(Range("D1:D20000"))
        q = Round(n / 2, 0)
    '    lineaMS linea1, Cells(2, 4), Cells(2, 5), Cells(2, 6), Cells(n, 4), Cells(n, 5), Cells(n, 6), 3, "0"
        x1 = Replace(Cells(2, 4), ",", "."): y1 = Replace(Cells(2, 5), ",", ".")
        x2 = Replace(Cells(q, 4), ",", "."): y2 = Replace(Cells(q, 5), ",", ".")
        x3 = Replace(Cells(n, 4), ",", "."): y3 = Replace(Cells(n, 5), ",", ".")
        Point1 = x1 & "," & y1:    Point2 = x2 & "," & y2:    Point3 = x3 & "," & y3
        AutoCAD.Application.ActiveDocument.SendCommand "_arc" & vbCr & Point1 & vbCr & Point2 & vbCr & Point3 & vbCr
        i = i + 1
        Poly.Delete
        Next
    ssetObj.Delete
    Range("A2:Z20000").Value = ""
End Sub

Public Function Seleccion(nombre As Variant, StartX As Variant, StartY As Variant, StartZ As Variant, EndX As Variant, EndY As Variant, EndZ As Variant) As AcadSelectionSet
    'Declaración de Variables
    Dim mode As Integer
    Dim PuntoA(0 To 2) As Double:    Dim PuntoB(0 To 2) As Double
    'Creando Selección
    Set nombre = AutoCAD.Application.ActiveDocument.SelectionSets.Add("SSE1")
    mode = acSelectionSetCrossing
    PuntoA(0) = StartX:     PuntoA(1) = StartY:         PuntoA(2) = StartZ
    PuntoB(0) = EndX:       PuntoB(1) = EndY:           PuntoB(2) = EndZ
    nombre.Select mode, PuntoA, PuntoB
End Function

Sub LargoLineas2()

Dim Poly As AcadLWPolyline
Dim Linea As AcadLine
Dim i As Long

Range("A2:Z20000").Value = ""
'activarAcad

Set ssetObj = AutoCAD.Application.ActiveDocument.SelectionSets.Add(Rnd)
ssetObj.SelectOnScreen

On Error Resume Next
For Each Poly In ssetObj
    
'EXTRAYENDO PUNTOS POLYLINEA
Bound = UBound(Poly.Coordinates)
i = 2
x = 0
Y = 1
For d = 0 To Bound / 2
Cells(i, 4) = Poly.Coordinates(x)
Cells(i, 4) = Replace(Cells(i, 4), ",", ".")
Cells(i, 4).NumberFormat = "0.0000"
Cells(i, 4).EntireColumn.AutoFit
Cells(i, 5) = Poly.Coordinates(Y)
Cells(i, 5) = Replace(Cells(i, 5), ",", ".")
Cells(i, 5).NumberFormat = "0.0000"
Cells(i, 5).EntireColumn.AutoFit
If Err Then Err.Clear
x = x + 2
Y = Y + 2
i = i + 1
Next d
        
        
'EXTRAYENDO PUNTOS POLYLINEA
r = 0
i = 2
For d = 0 To Bound
Cells(i, 6) = Poly.Coordinates(r)
If Err Then Err.Clear
i = i + 1
r = r + 1
Next d

NuevoLayer Layer0, "Temporal", 2, "Continuous", 0, True
Activalayer "Temporal"

Dim factor As Double
p = AutoCAD.Application.ActiveDocument.Utility.GetReal((i / 2) - 2 & " (Segmentos) Dividir en ?: ")
segmentos = (i / 2) - 2
f = Round(segmentos / p, 0)


ApagaTodosLosLayers

Dim layerObj As AcadLayer
Set layerObj = AutoCAD.Application.ActiveDocument.Layers.Item("Temporal")
layerObj.Freeze = False
layerObj.LayerOn = True

'        For c = 1 To p

For i = 2 To f
lineaMS Linea0, Cells(i, 4), Cells(i, 5), 0, Cells(i + 1, 4), Cells(i + 1, 5), 0, 256, "Temporal"
lineaMS Linea0, Cells(i + 1, 4), Cells(i + 1, 5), 0, Cells(i + 2, 4), Cells(i + 2, 5), 0, 256, "Temporal"
i = i + 1
f = f + f
Next

'        Next c




    ZoomExtents
    AutoCAD.Application.ActiveDocument.SendCommand "_ai_selall" & vbCr & "_j" & vbCr 'Ejecutando Comando Joint
    Next Poly
    ssetObj.Delete

    Range("A2:Z20000").Value = ""
    Dim n As Integer
    
    i = 2
    On Error Resume Next
    Set modelspace = AutoCAD.Application.ActiveDocument.modelspace
    For Each Entidad In modelspace
        If Entidad.layer = "Temporal" Then
        
        Dim Poly2 As AcadLWPolyline
        Set Poly2 = Entidad
            Bound2 = UBound(Poly2.Coordinates)
            x = 0
            Y = 1
            For d = 0 To Bound2 / 2
            Cells(i, 4) = Round(Poly2.Coordinates(x), 4)
            Cells(i, 4) = Replace(Cells(i, 4), ",", ".")
            Cells(i, 5) = Round(Poly2.Coordinates(Y), 4)
            If Err Then Err.Clear
            x = x + 2
            Y = Y + 2
            i = i + 1
            Next d

        n = Application.WorksheetFunction.CountA(Range("D1:D20000"))
        q = Round(n / 2, 0)
        x1 = Replace(Cells(2, 4), ",", "."): y1 = Replace(Cells(2, 5), ",", ".")
        x2 = Replace(Cells(q, 4), ",", "."): y2 = Replace(Cells(q, 5), ",", ".")
        x3 = Replace(Cells(n, 4), ",", "."): y3 = Replace(Cells(n, 5), ",", ".")
        Point1 = x1 & "," & y1:    Point2 = x2 & "," & y2:    Point3 = x3 & "," & y3
        AutoCAD.Application.ActiveDocument.SendCommand "_arc" & vbCr & Point1 & vbCr & Point2 & vbCr & Point3 & vbCr
        i = i + 1
        Poly2.Delete
        Range("A2:Z20000").Value = ""
        Else:
        End If
    Next Entidad


EnciendeTodosLosLayers
ZoomAll
    
End Sub



Sub prubea()

    Range("A2:Z20000").Value = ""
    Dim n As Integer
    
    i = 2
    On Error Resume Next
    Set modelspace = AutoCAD.Application.ActiveDocument.modelspace
    For Each Entidad In modelspace
        If Entidad.layer = "Temporal" Then
        
        Dim Poly2 As AcadLWPolyline
        Set Poly2 = Entidad
        

        
            Bound2 = UBound(Poly2.Coordinates)
            x = 0
            Y = 1
            For d = 0 To Bound2 / 2
            Cells(i, 4) = Round(Poly2.Coordinates(x), 4)
            Cells(i, 4) = Replace(Cells(i, 4), ",", ".")
            Cells(i, 5) = Round(Poly2.Coordinates(Y), 4)
            If Err Then Err.Clear
            x = x + 2
            Y = Y + 2
            i = i + 1
            Next d

        n = Application.WorksheetFunction.CountA(Range("D1:D20000"))
        q = Round(n / 2, 0)
        x1 = Replace(Cells(2, 4), ",", "."): y1 = Replace(Cells(2, 5), ",", ".")
        x2 = Replace(Cells(q, 4), ",", "."): y2 = Replace(Cells(q, 5), ",", ".")
        x3 = Replace(Cells(n, 4), ",", "."): y3 = Replace(Cells(n, 5), ",", ".")
        Point1 = x1 & "," & y1:    Point2 = x2 & "," & y2:    Point3 = x3 & "," & y3
        AutoCAD.Application.ActiveDocument.SendCommand "_arc" & vbCr & Point1 & vbCr & Point2 & vbCr & Point3 & vbCr
        i = i + 1
        
'        Poly.Delete
        'Range("A2:Z20000").Value = ""
        
        Else:
        End If
    Next Entidad

End Sub


