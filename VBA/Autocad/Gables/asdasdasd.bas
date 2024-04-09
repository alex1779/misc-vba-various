Attribute VB_Name = "asdasdasd"
Dim Entidad As AcadEntity
Dim i As Integer
Dim modelspace As AcadModelSpace
Dim PaperSpace As AcadPaperSpace
Dim ssetObj As AcadSelectionSet
Dim layerObj As AcadLayer
Dim LayerCol As AcadLayers
Dim Linea As AcadLine


Function activarAcad()
Dim objacadapp As AcadApplication
Set objacadapp = GetObject(, "Autocad.application")
AppActivate objacadapp.Caption
AutoCAD.WindowState = acMax
End Function

Public Function Format()
Range("A:z").NumberFormat = "0"
Range("A:D").EntireColumn.AutoFit
End Function
Sub Main()

BorrarDatosLibro
extrae_propiedades_modelspace
extrae_propiedades_paperspace
Extraccion_Datos
Extraccion_Datos_Lineas

End Sub

Public Function extrae_propiedades_modelspace()

NuevaHoja "Model Space"

Set modelspace = AutoCAD.Application.ActiveDocument.modelspace

    i = 1
    For Each Entidad In modelspace
    i = i + 1
    Cells(i, 1) = i - 1
    Cells(i, 2) = Entidad.ObjectName
    Cells(i, 3) = Entidad.ObjectID
    Next Entidad
    
    Decir "hay " & i & " entidades en el espacio modelo"
    
    Format
End Function

Public Function extrae_propiedades_paperspace()

NuevaHoja "Paper Space"

Set PaperSpace = AutoCAD.Application.ActiveDocument.PaperSpace
    i = 1
    For Each Entidad In PaperSpace
    i = i + 1
    Cells(i, 1) = i - 1
    Cells(i, 2) = Entidad.ObjectName
    Cells(i, 3) = Entidad.ObjectID
    Next Entidad
    
    Decir "hay " & i & " entidades en el paper space"
    
    Format
End Function

Sub extrae_propiedades_seleccion()

Set ssetObj = AutoCAD.Application.ActiveDocument.SelectionSets.Add("SSE1")
ssetObj.SelectOnScreen
cantidad_objetos = ssetObj.Count
i = 1
    For Each Entidad In ssetObj
    i = i + 1
    Cells(i, 1) = i - 1
    Cells(i, 2) = Entidad.ObjectName
    Cells(i, 3) = Entidad.ObjectID

    Next Entidad
ssetObj.Delete
End Sub

Public Function Crear_Hojas_Entidades()

Set modelspace = AutoCAD.Application.ActiveDocument.modelspace
Set PaperSpace = AutoCAD.Application.ActiveDocument.PaperSpace
Set LayersCollection = AutoCAD.Application.ActiveDocument.Layers

'ENTIDADES
For Each Entidad In modelspace
NuevaHoja Entidad.ObjectName & "-MS"
Next Entidad

For Each Entidad In PaperSpace
NuevaHoja Entidad.ObjectName & "-PS"
Next Entidad


'LAYERS
For Each layerObj In LayersCollection
NuevaHoja layerObj.ObjectName '& "-MS"
Next layerObj



End Function


Public Function Extraccion_Datos()
Dim t As Single
t = Timer

Crear_Hojas_Entidades


'ENTIDADES GENERALES
Set modelspace = AutoCAD.Application.ActiveDocument.modelspace
For Each Entidad In modelspace
Sheets(Entidad.ObjectName & "-MS").Select
i = Application.WorksheetFunction.CountA(Range("A2:A2000")) + 2
Cells(i, 1) = i - 1
Cells(i, 2) = Entidad.ObjectName
Cells(i, 3) = Entidad.ObjectID
Format
Next Entidad

Set PaperSpace = AutoCAD.Application.ActiveDocument.PaperSpace
For Each Entidad In PaperSpace
Sheets(Entidad.ObjectName & "-PS").Select
i = Application.WorksheetFunction.CountA(Range("A2:A2000")) + 2
Cells(i, 1) = i - 1
Cells(i, 2) = Entidad.ObjectName
Cells(i, 3) = Entidad.ObjectID
Format
Next Entidad


'LAYERS
Set LayersCollection = AutoCAD.Application.ActiveDocument.Layers
For Each layerObj In LayersCollection
Sheets(layerObj.ObjectName).Select
i = Application.WorksheetFunction.CountA(Range("A2:A2000")) + 2
Cells(i, 1) = i - 1
Cells(i, 2) = layerObj.ObjectName
Cells(i, 3) = layerObj.Name
Cells(i, 4) = layerObj.Color
Cells(i, 5) = layerObj.LineType
Cells(i, 6) = layerObj.LineWeight
Cells(i, 7) = layerObj.Plottable
Format
Next layerObj



BorrarHojasVacias
tiempo = Round(Timer - t)
Application.Speech.Speak tiempo & " segundos", True, , True
End Function


Public Function Extraccion_Datos_Lineas()
Dim Linea As AcadLine
Set modelspace = AutoCAD.Application.ActiveDocument.modelspace
i = 2
For Each Entidad In modelspace
    If Entidad.ObjectName = "AcDbLine" Then
    Sheets(Entidad.ObjectName & "-MS").Select
    Set Linea = Entidad
    Cells(i, 4) = Linea.StartPoint(0)
    Cells(i, 5) = Linea.StartPoint(1)
    Cells(i, 6) = Linea.StartPoint(2)
    Cells(i, 7) = Linea.EndPoint(0)
    Cells(i, 8) = Linea.EndPoint(1)
    Cells(i, 9) = Linea.EndPoint(2)
    Cells(i, 10) = Linea.Color
    Cells(i, 11) = Linea.layer
    Range("D:I").NumberFormat = "0.0000000"
    Range("D:I").EntireColumn.AutoFit
    Range("J:K").NumberFormat = "0"
    Range("J:K").EntireColumn.AutoFit
    i = i + 1
    Else:
    End If
Next Entidad


Set PaperSpace = AutoCAD.Application.ActiveDocument.PaperSpace
i = 2
For Each Entidad In PaperSpace
    If Entidad.ObjectName = "AcDbLine" Then
    Sheets(Entidad.ObjectName & "-PS").Select
    Set Linea = Entidad
    Cells(i, 4) = Linea.StartPoint(0)
    Cells(i, 5) = Linea.StartPoint(1)
    Cells(i, 6) = Linea.StartPoint(2)
    Cells(i, 7) = Linea.EndPoint(0)
    Cells(i, 8) = Linea.EndPoint(1)
    Cells(i, 9) = Linea.EndPoint(2)
    Cells(i, 10) = Linea.Color
    Cells(i, 11) = Linea.layer
    Range("D:I").NumberFormat = "0.0000000"
    Range("D:I").EntireColumn.AutoFit
    Range("J:K").NumberFormat = "0"
    Range("J:K").EntireColumn.AutoFit
    i = i + 1
    Else:
    End If
Next Entidad

End Function
Public Function punto(nombre As Variant, CenterX As Variant, CenterY As Variant, CenterZ As Variant, Color As AcColor, layer As String)
    'Declaración de Variables
    Dim Center(0 To 2) As Double
    Center(0) = CenterX:     Center(1) = CenterY:         Center(2) = CenterZ
    'Creando nuevo punto
    Set nombre = AutoCAD.Application.ActiveDocument.modelspace.AddPoint(Center)
    nombre.Color = Color
    nombre.layer = layer
End Function
Sub ColocarPuntos()

i = Application.WorksheetFunction.CountA(Range("A2:A10000"))

For n = 1 To i

punto Punto1, Cells(n + 1, 4), Cells(n + 1, 5), Cells(n + 1, 6), acGreen, "0"
    AutoCAD.Application.ActiveDocument.SendCommand "regen" & vbCr 'Regen


Next

Dim Spiline As AcadSpline
Set Spiline = AutoCAD.Application.ActiveDocument.modelspace.AddSpline(PointsArray, StartTangent, EndTangent)


End Function


Sub reemplaza_lineas_seleccion()

    Dim Minpoint(0 To 2) As Double
    Dim MaxPoint(0 To 2) As Double
    Dim comienza As Double
    activarAcad
    
    Set modelspace = AutoCAD.Application.ActiveDocument.modelspace

    NuevoLayer Temporal, "Temporal", 2, "Continuous", 0, True
    Dim newlayer As AcadLayer
    Set newlayer = AutoCAD.Application.ActiveDocument.Layers.Add("Temporal")
    Activalayer "temporal"

    Set ssetObj = AutoCAD.Application.ActiveDocument.SelectionSets.Add("SSE36")
    ssetObj.SelectOnScreen
    i = 1
    For Each Linea In ssetObj
    Linea.layer = "temporal"
    Linea.Color = 256
    Next Linea
    ssetObj.Delete

    ApagaLayer "0"
    ZoomExtents
    AutoCAD.Application.ActiveDocument.SendCommand "_ai_selall" & vbCr & "_j" & vbCr 'Ejecutando Comando Joint

    AutoCAD.Application.ActiveDocument.SendCommand "regen" & vbCr 'Regen


    Set ssetObj = AutoCAD.Application.ActiveDocument.SelectionSets.Add("SSE36")
    ssetObj.SelectOnScreen
    i = 1
    
    Dim pollinea As AcadLWPolyline
    
    For Each Entidad In ssetObj
    

    
    If Entidad.ObjectName = "AcDbPolyline" Then

        Set pollinea = Entidad

        MsgBox pollinea.Length
'        MsgBox pollinea.StartPoint(1)
'        MsgBox pollinea.StartPoint(2)

    Else:
    End If
    Next Entidad
    ssetObj.Delete
        

    EnciendeLayer "0"

End Sub
    



Sub Convertir_arco()


        activarAcad
    Dim Linea As AcadLine
    Set modelspace = AutoCAD.Application.ActiveDocument.modelspace
    i = 2
    Set ssetObj = AutoCAD.Application.ActiveDocument.SelectionSets.Add("SSE561")
    ssetObj.SelectOnScreen
    
    For Each Entidad In ssetObj
    If Entidad.ObjectName = "AcDbLine" Then
    Sheets(Entidad.ObjectName & "-MS").Select
    Set Linea = Entidad
    Cells(i, 4) = Linea.StartPoint(0)
    Cells(i, 5) = Linea.StartPoint(1)
    Cells(i, 6) = Linea.StartPoint(2)
    Cells(i, 7) = Linea.EndPoint(0)
    Cells(i, 8) = Linea.EndPoint(1)
    Cells(i, 9) = Linea.EndPoint(2)
    
    'Entidad.Delete

    i = i + 1
    Else:
    End If
    Next Entidad
    ssetObj.Delete
    
    
    For a = 2 To i - 1
    For b = 2 To i - 1

    lineaMS linea1, Cells(a, 4), Cells(a, 5), Cells(a, 6), Cells(b, 7), Cells(b, 8), Cells(b, 9), 3, "0"
    
    Next b
    
    Next a

    For a = 2 To i
    Cells(a, 4) = ""
    Cells(a, 5) = ""
    Cells(a, 6) = ""
    Cells(a, 7) = ""
    Cells(a, 8) = ""
    Cells(a, 9) = ""
    Next a
eligemasgrande
End Sub


Sub eligemasgrande()

    activarAcad
    Dim Linea As AcadLine
    Set modelspace = AutoCAD.Application.ActiveDocument.modelspace
    
    Set ssetObj = AutoCAD.Application.ActiveDocument.SelectionSets.Add("SSE561")
    ssetObj.SelectOnScreen
    i = 2
    On Error Resume Next
    For Each Linea In ssetObj
    If Linea.Color = acGreen Then
    Cells(i, 4) = Linea.Length
    Else:
    End If
    i = i + 1
    Next Linea
    
    For Each Linea In ssetObj
    If Linea.Color = acGreen Then
        If Linea.Length = WorksheetFunction.Max(Range(Cells(2, 4), Cells(i, 4))) Then
        Else:
        Linea.Delete
        End If
    Else:
    End If
    Next Linea
    
    x = 2
    For x = 2 To i
    Cells(x, 4) = ""
    x = x + 1
    Next x
    
    AutoCAD.Application.ActiveDocument.SendCommand "regen" & vbCr 'Regen
    ssetObj.Delete

End Sub


Sub ReduccionNodos()
    Range("A2:Z20000").Value = ""
    activarAcad
    Dim Linea As AcadLine
    Set modelspace = AutoCAD.Application.ActiveDocument.modelspace
    i = 2
    Set ssetObj = AutoCAD.Application.ActiveDocument.SelectionSets.Add("SSE562")
    ssetObj.SelectOnScreen
    For Each Linea In ssetObj
'    If Entidad.ObjectName = "AcDbLine" Then
'    Set linea = Entidad
    Cells(i, 4) = Linea.StartPoint(0)
    Cells(i, 5) = Linea.StartPoint(1)
    Cells(i, 6) = Linea.StartPoint(2)
    Cells(i, 7) = Linea.EndPoint(0)
    Cells(i, 8) = Linea.EndPoint(1)
    Cells(i, 9) = Linea.EndPoint(2)
    Linea.Delete
    i = i + 1
'    Else:
'    End If
    Next Linea
    ssetObj.Delete
    
    Dim factor As Double
    factor = AutoCAD.Application.ActiveDocument.Utility.GetReal("Ingrese Factor: ")
    
    
    For b = 2 To i

    lineaMS linea1, Cells(b, 4), Cells(b, 5), Cells(b, 6), Cells(b + factor, 4), Cells(b + factor, 5), Cells(b + factor, 6), 255, "0"
    b = b + (factor - 1)

    Next b
    

End Sub


Sub repite()


For i = 1 To 5
ConvSplineToArc
Next


End Sub




Sub ConvSplineToArc()
    Range("A2:Z20000").Value = ""
    'activarAcad
    Set modelspace = AutoCAD.Application.ActiveDocument.modelspace
    
    Dim Spline As AcadSpline
    i = 2
    Set ssetObj = AutoCAD.Application.ActiveDocument.SelectionSets.Add("SSE562")
    ssetObj.SelectOnScreen
    On Error Resume Next
    For Each Spline In ssetObj
    p = Spline.NumberOfFitPoints - 1
    Cells(i, 4) = Spline.GetFitPoint(0)(0)
    Cells(i, 5) = Spline.GetFitPoint(0)(1)
    Cells(i, 6) = Spline.GetFitPoint(0)(2)
    Cells(i, 7) = Spline.GetFitPoint(p)(0)
    Cells(i, 8) = Spline.GetFitPoint(p)(1)
    Cells(i, 9) = Spline.GetFitPoint(p)(2)
    'lineaMS linea1, Cells(2, 4), Cells(2, 5), Cells(2, 6), Cells(2, 7), Cells(2, 8), Cells(2, 9), 3, "0"
    i = i + 1
    Next
    ssetObj.Delete
    AutoCAD.Application.ActiveDocument.SendCommand "SPLINEDIT" & vbCr & "p" & vbCr & "p" & vbCr & "30" & vbCr

    Dim ptCoord As Double
    Set ssetObj = AutoCAD.Application.ActiveDocument.SelectionSets.Add("SSE569")
    ssetObj.SelectOnScreen




End Sub
