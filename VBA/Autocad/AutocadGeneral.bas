Attribute VB_Name = "AutocadGeneral"
Public Enum vistas
top
Bottom
Left
Right
Front
Back
SW_Isometric
SE_Isometric
NE_Isometric
NW_Isometric
End Enum

Public Enum Sombras
wireframe2d
Wireframe
Hidden
Realistic
Conceptual
shaded
shaded_with_Edges
shades_of_Gray
SKetchy
X_ray
End Enum

Public Enum colors
acByBlock = 7
acRed = 1
acYellow = 2
acGreen = 3
acCyan = 4
acBlue = 5
acMagenta = 6
acWhite = 7
End Enum

Public Function vista(vist As vistas)
Select Case vist
Case vistas.top
view = "_Top"
Case vistas.Bottom
view = "_Bottom"
Case vistas.Left
view = "_Left"
Case vistas.Right
view = "_Right"
Case vistas.Front
view = "_Front"
Case vistas.Back
view = "_Back"
Case vistas.SW_Isometric
view = "_swiso"
Case vistas.SE_Isometric
view = "_seiso"
Case vistas.NE_Isometric
view = "_neiso"
Case vistas.NW_Isometric
view = "_nwiso"
End Select
AutoCAD.Application.ActiveDocument.SendCommand "_-view" & vbCr & view & vbCr
End Function


Public Function Sombreado(sombra As Sombras)
Select Case sombra
Case Sombras.wireframe2d
view = "2"
Case Sombras.Wireframe
view = "W"
Case Sombras.Hidden
view = "H"
Case Sombras.Realistic
view = "_R"
Case Sombras.Conceptual
view = "_C"
Case Sombras.shaded_with_Edges
view = "S"
Case Sombras.shades_of_Gray
view = "G"
Case Sombras.SKetchy
view = "SK"
Case Sombras.X_ray
view = "X"
End Select
AutoCAD.Application.ActiveDocument.SendCommand "shademode" & vbCr & view & vbCr
End Function







' LAYER'S

Public Function NuevoLayer(nombre As Variant, Name As Variant, Color As colors, linetype As String, LineWeight As AcLineWeight, Plottable As Boolean)
    Set nombre = AutoCAD.Application.ActiveDocument.Layers.Add(Name)
    nombre.Color = Color
    nombre.linetype = linetype
    nombre.LineWeight = LineWeight
    nombre.Plottable = Plottable
    nombre = nombre.Name
End Function


Public Function Activalayer(nombre As Variant)
    Dim layername As AcadLayer
    Set layername = AutoCAD.Application.ActiveDocument.Layers.Item(nombre)
    AutoCAD.Application.ActiveDocument.ActiveLayer = layername
End Function
Public Function ApagaLayer(nombre As Variant)
    On Error Resume Next
    Dim layerObj As AcadLayer
    Set layerObj = AutoCAD.Application.ActiveDocument.Layers.Item(nombre)
    layerObj.Freeze = True
    layerObj.LayerOn = False
End Function
Public Function EnciendeLayer(nombre As Variant)
    On Error Resume Next
    Dim layerObj As AcadLayer
    Set layerObj = AutoCAD.Application.ActiveDocument.Layers.Item(nombre)
    layerObj.Freeze = False
    layerObj.LayerOn = True
    
End Function



Public Function EnciendeTodosLosLayers()
    On Error Resume Next
    Dim layerObj As AcadLayer
    For Each layerObj In AutoCAD.Application.ActiveDocument.Layers
    layerObj.Freeze = False
    layerObj.LayerOn = True
    Next layerObj
End Function
Public Function ApagaTodosLosLayers()
    On Error Resume Next
    Dim layerObj As AcadLayer
    For Each layerObj In AutoCAD.Application.ActiveDocument.Layers
    layerObj.Freeze = True
    layerObj.LayerOn = False
    Next layerObj
End Function

Public Function BorrarTodo()    'Selecciona Todo y Borra
    entidades = AutoCAD.Application.ActiveDocument.ModelSpace.Count
    If entidades > 0 Then
        AutoCAD.Application.ActiveDocument.SendCommand "_ai_selall" & vbCr & "_.erase" & vbCr
    Else:
    End If
End Function
Public Function ActivarModelSpace()    'Activando: Model Space
    AutoCAD.Application.ActiveDocument.ActiveSpace = acModelSpace
End Function
Public Function ActivarPaperSpace()    'Activando: Paper Space
    AutoCAD.Application.ActiveDocument.ActiveSpace = acPaperSpace
End Function

Public Function CopiarTodoPunto0()   'Copia Todo desde el punto 0,0,0
    AutoCAD.Application.ActiveDocument.SendCommand "_ai_selall" & vbCr & "_copybase" & vbCr & "0,0,0 " & vbCr
End Function
Public Function PegarPunto0()   'Pega Todo desde el punto 0,0,0
    AutoCAD.Application.ActiveDocument.SendCommand "_pasteclip " & vbCr & "0,0,0 " & vbCr
End Function
Public Function Pegar(Punto As Variant)   'Pega Todo desde el punto 0,0,0
    AutoCAD.Application.ActiveDocument.SendCommand "_pasteclip " & vbCr & Punto & vbCr
End Function



Public Function NuevoSolido(ByVal nombre As Variant, ByVal Punto1 As Variant, ByVal Punto2 As Variant, ByVal Punto3 As Variant, ByVal Punto4 As Variant, ByVal Punto5 As Variant, ByVal Punto6 As Variant, ByVal Punto7 As Variant, ByVal Punto8 As Variant, ByVal Punto9 As Variant, ByVal Punto10 As Variant, ByVal Punto11 As Variant, ByVal Punto12 As Variant, ByVal ThkRoof As Variant) As AcadSolid


    Dim Point1(0 To 2) As Double
    Dim Point2(0 To 2) As Double
    Dim Point3(0 To 2) As Double
    Dim Point4(0 To 2) As Double
    Dim curves(0 To 3) As AcadEntity
    Dim objetoregionTecho As Variant
    Dim objetosolidoTecho As Acad3DSolid
    
    'Definiendo los puntos de Techo
    Point1(0) = Punto1:     Point1(1) = Punto2:     Point1(2) = Punto3
    Point2(0) = Punto4:     Point2(1) = Punto5:     Point2(2) = Punto6
    Point3(0) = Punto7:     Point3(1) = Punto8:     Point3(2) = Punto9
    Point4(0) = Punto10:    Point4(1) = Punto11:    Point4(2) = Punto12
    
    'Definiendo las lineas de Techo
    Set curves(0) = AutoCAD.Application.ActiveDocument.ModelSpace.AddLine(Point1, Point2)
    Set curves(1) = AutoCAD.Application.ActiveDocument.ModelSpace.AddLine(Point2, Point3)
    Set curves(2) = AutoCAD.Application.ActiveDocument.ModelSpace.AddLine(Point3, Point4)
    Set curves(3) = AutoCAD.Application.ActiveDocument.ModelSpace.AddLine(Point4, Point1)
    objetoregionTecho = AutoCAD.Application.ActiveDocument.ModelSpace.AddRegion(curves) 'Creando la región de Techo
    'Eliminando las lineas de Piso
    curves(0).Delete
    curves(1).Delete
    curves(2).Delete
    curves(3).Delete
    Set objetosolidoTecho = AutoCAD.Application.ActiveDocument.ModelSpace.AddExtrudedSolid(objetoregionTecho(0), ThkRoof, 0) 'Creando el solido de Techo
    objetoregionTecho(0).Delete 'Eliminando la región de Techo
    ZoomExtents
End Function






Public Function Linea(nombre As Variant, StartX As Variant, StartY As Variant, StartZ As Variant, EndX As Variant, EndY As Variant, EndZ As Variant, Color As AcColor, Optional layer As Variant)
    'Declaración de Variables
    Dim PuntoA(0 To 2) As Double:    Dim PuntoB(0 To 2) As Double
    PuntoA(0) = StartX:     PuntoA(1) = StartY:         PuntoA(2) = StartZ
    PuntoB(0) = EndX:       PuntoB(1) = EndY:           PuntoB(2) = EndZ
    'Creando nueva linea
    Set nombre = AutoCAD.Application.ActiveDocument.ModelSpace.AddLine(PuntoA, PuntoB)
    nombre.Color = Color
    On Error Resume Next
    nombre.layer = layer
End Function
Public Function Circulo(nombre As Variant, CenterX As Variant, CenterY As Variant, CenterZ As Variant, Radious As Variant, Color As AcColor, layer As String)
    'Declaración de Variables
    ''Dim Circulo As AcadCircle
    Dim center(0 To 2) As Double
    center(0) = CenterX:     center(1) = CenterY:         center(2) = CenterZ

    'Creando nuevo círculo
    Set nombre = AutoCAD.Application.ActiveDocument.ModelSpace.AddCircle(center, Radious)
    nombre.Color = Color
    nombre.layer = layer
End Function
Public Function texto(Space As AcActiveSpace, nombre As Variant, PuntoAX As Variant, PuntoAY As Variant, PuntoAZ As Variant, Valor As String, Altura As Double, alineacion As AcAlignment, estilo As String, layer As String, Ancho As Variant)

    'Iniciación de Variables
    Dim Punto0(0 To 2) As Double:    Dim PuntoA(0 To 2) As Double
    Punto0(0) = 0: Punto0(1) = 0: Punto0(2) = 0
    PuntoA(0) = PuntoAX: PuntoA(1) = PuntoAY: PuntoA(2) = PuntoAZ
    'Creando nueva Texto
    If Space = 1 Then
    Set nombre = AutoCAD.Application.ActiveDocument.ModelSpace.AddText(Valor, Punto0, Altura)
    Else:
    Set nombre = AutoCAD.Application.ActiveDocument.PaperSpace.AddText(Valor, Punto0, Altura)
    End If
    nombre.Alignment = alineacion
    nombre.TextAlignmentPoint = PuntoA
    nombre.StyleName = estilo
    nombre.layer = layer
    nombre.ScaleFactor = Ancho
    
End Function

Public Function Punto(nombre As Variant, CenterX As Variant, CenterY As Variant, CenterZ As Variant, Color As AcColor, layer As String)
    'Declaración de Variables
    Dim center(0 To 2) As Double
    center(0) = CenterX:     center(1) = CenterY:         center(2) = CenterZ
    'Creando nuevo punto
    Set nombre = AutoCAD.Application.ActiveDocument.ModelSpace.AddPoint(center)
    nombre.Color = Color
    nombre.layer = layer
End Function
Public Function Seleccion(nombre As Variant, StartX As Variant, StartY As Variant, StartZ As Variant, EndX As Variant, EndY As Variant, EndZ As Variant) As AcadSelectionSet
    'Dim ssetObj As AcadSelectionSet
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

Public Function LineaPaperSpace(nombre As Variant, StartX As Variant, StartY As Variant, StartZ As Variant, EndX As Variant, EndY As Variant, EndZ As Variant, Color As AcColor, layer As String)
    'Declaración de Variables
    Dim PuntoA(0 To 2) As Double:    Dim PuntoB(0 To 2) As Double
    PuntoA(0) = StartX:     PuntoA(1) = StartY:         PuntoA(2) = StartZ
    PuntoB(0) = EndX:       PuntoB(1) = EndY:           PuntoB(2) = EndZ
    'Creando nueva linea
    Set nombre = AutoCAD.Application.ActiveDocument.PaperSpace.AddLine(PuntoA, PuntoB)
    nombre.Color = Color
    nombre.layer = layer
End Function
Public Function NuevoDibujo()   'Nuevo dibujo en autocad
        AutoCAD.Application.Documents.Add (acad)    'New Drawing
End Function
Public Function Overkill()   'Overkill
    AutoCAD.Application.ActiveDocument.SendCommand "-overkill" & vbCr & "all" & vbCr & vbCr & vbCr  'Overkill from all entities
End Function
Public Function Cerrar(cambios As Boolean)   'Close Acad File
        AutoCAD.Application.ActiveDocument.Close cambios
End Function

Public Function CerrarTodos()   'Close Acad File

    df = AutoCAD.Application.Documents.Count
    For i = 1 To df
    AutoCAD.Application.ActiveDocument.Close False
    Next i
End Function
Public Function BorrarTodos(nombre As AcEntityName)

    If nombre = 1 Then ent = "AcDbFace"



    Dim Objeto As AcadEntity
    For Each Objeto In AutoCAD.Application.ActiveDocument.ModelSpace
        If Objeto.ObjectName = ent Then
        Objeto.Delete
        End If
    Next Objeto
        AutoCAD.Application.ActiveDocument.SendCommand "regen" & vbCr
End Function

Public Function Borrar(opcion As String)
    If opcion = "lineas" Then Var = "AcDbLine"
    If opcion = "solidos" Then Var = "AcDb3dSolid"
    If opcion = "puntos" Then Var = "AcDbPoint"
    If opcion = "arcos" Then Var = "AcDbArc"
    If opcion = "polilineas" Then Var = "AcDbPolyline"
    If opcion = "bloques" Then Var = "AcDbBlockReference"
    If opcion = "viewports" Then Var = "AcDbViewport"
    If opcion = "circulos" Then Var = "AcDbCircle"
    If opcion = "elipse" Then Var = "AcDbEllipse"
    If opcion = "spline" Then Var = "AcDbSpline"
    If opcion = "3dface" Then Var = "AcDbFace"
    If opcion = "Region" Then Var = "AcDbRegion"

    Dim Objeto As AcadEntity
    For Each Objeto In AutoCAD.Application.ActiveDocument.ModelSpace
        If Objeto.ObjectName = Var Then
        Objeto.Delete
        End If
    Next Objeto
        AutoCAD.Application.ActiveDocument.SendCommand "regen" & vbCr
End Function

Public Function BorrarExepto(opcion As String)
    If opcion = "lineas" Then Var = "AcDbLine"
    If opcion = "solidos" Then Var = "AcDb3dSolid"
    If opcion = "puntos" Then Var = "AcDbPoint"
    If opcion = "arcos" Then Var = "AcDbArc"
    If opcion = "polilineas" Then Var = "AcDbPolyline"
    If opcion = "bloques" Then Var = "AcDbBlockReference"
    If opcion = "viewports" Then Var = "AcDbViewport"
    If opcion = "circulos" Then Var = "AcDbCircle"
    If opcion = "elipse" Then Var = "AcDbEllipse"
    If opcion = "spline" Then Var = "AcDbSpline"
    Dim Objeto As AcadEntity
    For Each Objeto In AutoCAD.Application.ActiveDocument.ModelSpace
        If Objeto.ObjectName = Var Then
        Else
        Objeto.Delete
        End If
    Next Objeto
        AutoCAD.Application.ActiveDocument.SendCommand "regen" & vbCr
End Function

Public Function UnionSolidos()
    AutoCAD.Application.ActiveDocument.SendCommand "_ai_selall" & vbCr & "_union" & vbCr & "all" & vbCr & vbCr 'Union from all solids
End Function
Public Function Solprof()
    ActivarPaperSpace
    AutoCAD.Application.ActiveDocument.SendCommand "ms" & vbCr  'Activando: Viewport
    AutoCAD.Application.ActiveDocument.SendCommand "solprof" & vbCr & "all" & vbCr & vbCr & vbCr & vbCr & vbCr 'Solprof
End Function
Public Function ExplotaTodosBloques()
    Dim Entidad As AcadEntity
    Dim bloques As AcadBlockReference
    i = 1
    Do While i > 0
    For Each Entidad In AutoCAD.Application.ActiveDocument.ModelSpace
    If Entidad.ObjectName = "AcDbBlockReference" Then
    Set bloques = Entidad
    bloques.Explode
    Entidad.Delete
    End If
    Next Entidad
    i = 0
    For Each Entidad In AutoCAD.Application.ActiveDocument.ModelSpace
    If Entidad.ObjectName = "AcDbBlockReference" Then
    i = 1
    End If
    Next Entidad
    For Each Entidad In AutoCAD.Application.ActiveDocument.ModelSpace
    If Entidad.ObjectName = "AcDbBlockReference" Then
    i = i + 1
    End If
    Next Entidad
    Loop
    AutoCAD.Application.ActiveDocument.PurgeAll 'Purgando
End Function

Public Function UCSWorld()
    AutoCAD.Application.ActiveDocument.SendCommand "ucs" & vbCr & "w" & vbCr 'UCS
End Function

Public Function ucs(ucs2 As AcCoordinateSystem)
    AutoCAD.Application.ActiveDocument.SendCommand "ucs" & vbCr & ucs2 & vbCr 'UCS
End Function

Function activarAcad()
Dim objacadapp As AcadApplication
Set objacadapp = GetObject(, "Autocad.application")
AppActivate objacadapp.Caption
AutoCAD.WindowState = acMax
End Function
Public Function Regen()
    AutoCAD.Application.ActiveDocument.SendCommand "regen" & vbCr 'Regen

End Function

Public Sub LayersRotulo()
On Error Resume Next
    AutoCAD.Application.ActiveDocument.linetypes.Load "CENTER", "acad.lin"
    AutoCAD.Application.ActiveDocument.linetypes.Load "HIDDEN", "acad.lin"
    layer LayerObj1, ".3D MODEL", acGreen, "Continuous", acLnWt060, True
    layer LayerObj1, ".DIMENSION", acMagenta, "Continuous", acLnWt025, True
    layer LayerObj2, ".GENERAL ARRANGEMENT", acYellow, "Continuous", acLnWt040, True
    layer LayerObj3, ".HIDDEN", 8, "HIDDEN", acLnWt000, True
    layer LayerObj4, ".OUTLINE 1", acCyan, "Continuous", acLnWt060, True
    layer LayerObj5, ".OUTLINE 3", acRed, "Continuous", acLnWt050, True
    layer LayerObj6, ".TEXT", 50, "Continuous", acLnWt030, True
    layer LayerObj7, ".TITLE BLOCK ATTRIBUTES", acYellow, "Continuous", acLnWtByLwDefault, True
    layer LayerObj8, ".TITLE BLOCK OUTLINE", 151, "Continuous", acLnWtByLwDefault, True
    layer LayerObj9, ".TITLE BLOCK TEXT", acGreen, "Continuous", acLnWtByLwDefault, True
    layer LayerObj10, ".TITLE PAGE OUTLINE", 252, "Continuous", acLnWtByLwDefault, False
    layer LayerObj11, "0", acWhite, "Continuous", acLnWtByLwDefault, True
    layer LayerObj12, ".CENTER LINE", 8, "CENTER", acLnWtByLwDefault, True
    layer LayerObj13, "Defpoints", acWhite, "Continuous", acLnWtByLwDefault, False
End Sub
Public Function Caratula()
Proyecto = Application.InputBox("Nombre de Proyecto")
Dim linea1 As Variant
Dim n As Integer
n = Application.WorksheetFunction.CountA(Range("A1:A2000"))
Linea linea1, 0, 0, 0, 210, 0, 0, acByLayer, "0"
Linea linea1, 210, 0, 0, 210, 297, 0, acByLayer, "0"
Linea linea1, 210, 297, 0, 0, 297, 0, acByLayer, "0"
Linea linea1, 0, 297, 0, 0, 0, 0, acByLayer, "0"
ZoomExtents
Linea linea1, 20, 5, 0, 200, 5, 0, acByLayer, "0"
Linea linea1, 20, 15, 0, 200, 15, 0, acByLayer, "0"
Linea linea1, 20, 265, 0, 200, 265, 0, acByLayer, "0"
Linea linea1, 20, 271, 0, 200, 271, 0, acByLayer, "0"
Linea linea1, 200, 277, 0, 20, 277, 0, acByLayer, "0"
Linea linea1, 200, 5, 0, 200, 277, 0, acByLayer, "0"
Linea linea1, 20, 277, 0, 20, 5, 0, acByLayer, "0"
texto acModelSpace, text1, 25, 10, 0, "Autor: Alejandro Maggioni", 3, acAlignmentMiddleLeft, "Standard", "0", 1
texto acModelSpace, text1, 125, 10, 0, "Contact to: Alex1779@hotmail.com", 3, acAlignmentMiddleLeft, "Standard", "0", 1
texto acModelSpace, text1, 25, 274, 0, "PROYECT: " & Proyecto, 3, acAlignmentMiddleLeft, "Standard", "0", 1
texto acModelSpace, text1, 25, 268, 0, "TITLE: ", 3, acAlignmentMiddleLeft, "Standard", "0", 1
texto acModelSpace, text1, 160, 274, 0, "DATE: " & Date, 3, acAlignmentMiddleLeft, "Standard", "0", 1
'texto acModelSpace, text1, 160, 268, 0, "POINTS: " & n - 2, 3, acAlignmentMiddleLeft, "Standard", "0", 1
AutoCAD.Application.ActiveDocument.PurgeAll 'Purge
AutoCAD.Application.ActiveDocument.SendCommand "regen" & vbCr 'Regen
End Function




