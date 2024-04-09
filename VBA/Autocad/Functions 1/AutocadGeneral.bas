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
Public Function VistaNuevo(vista As vistas)
Select Case vista
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

Public Function NuevoLayer(nombre As Variant, Name As Variant, color As Variant, LineType As String, LineWeight As Variant, Plottable As Boolean)
    Set nombre = AutoCAD.Application.ActiveDocument.Layers.Add(Name)
    nombre.color = color
    nombre.LineType = LineType
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
    Dim layerObj As AcadLayer
    Set layerObj = AutoCAD.Application.ActiveDocument.Layers.Item(nombre)
    layerObj.Freeze = True
End Function
Public Function EnciendeLayer(nombre As Variant)
    Dim layerObj As AcadLayer
    Set layerObj = AutoCAD.Application.ActiveDocument.Layers.Item(nombre)
    layerObj.Freeze = False
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
Public Function Pegar(Punto As String)   'Pega Todo desde el punto 0,0,0
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






Public Function Linea(nombre As Variant, StartX As Variant, StartY As Variant, StartZ As Variant, EndX As Variant, EndY As Variant, EndZ As Variant, color As Variant, layer As Variant)
    'Declaración de Variables
    Dim PuntoA(0 To 2) As Double:    Dim PuntoB(0 To 2) As Double
    PuntoA(0) = StartX:     PuntoA(1) = StartY:         PuntoA(2) = StartZ
    PuntoB(0) = EndX:       PuntoB(1) = EndY:           PuntoB(2) = EndZ
    'Creando nueva linea
    Set nombre = AutoCAD.Application.ActiveDocument.ModelSpace.AddLine(PuntoA, PuntoB)
    nombre.color = color
    nombre.layer = layer
End Function
Public Function Circulo(nombre As Variant, CenterX As Variant, CenterY As Variant, CenterZ As Variant, Radious As Variant, color As Variant, layer As String)
    'Declaración de Variables
    ''Dim Circulo As AcadCircle
    Dim Center(0 To 2) As Double
    Center(0) = CenterX:     Center(1) = CenterY:         Center(2) = CenterZ

    'Creando nuevo círculo
    Set nombre = AutoCAD.Application.ActiveDocument.ModelSpace.AddCircle(Center, Radious)
    nombre.color = color
    nombre.layer = layer
End Function
Public Function texto(nombre As Variant, PuntoAX As Variant, PuntoAY As Variant, PuntoAZ As Variant, valor As String, Altura As Double, alineacion As Variant, estilo As String, layer As String, Ancho As Variant)
    'Iniciación de Variables
    Dim Punto0(0 To 2) As Double:    Dim PuntoA(0 To 2) As Double
    Punto0(0) = 0: Punto0(1) = 0: Punto0(2) = 0
    PuntoA(0) = PuntoAX: PuntoA(1) = PuntoAY: PuntoA(2) = PuntoAZ
    'Creando nueva Texto
    Set nombre = AutoCAD.Application.ActiveDocument.PaperSpace.AddText(valor, Punto0, Altura)
    nombre.Alignment = alineacion
    nombre.TextAlignmentPoint = PuntoA
    nombre.StyleName = estilo
    nombre.layer = layer
    nombre.ScaleFactor = Ancho
End Function

Public Function Punto(nombre As Variant, CenterX As Variant, CenterY As Variant, CenterZ As Variant, color As Variant, layer As String)
    'Declaración de Variables
    Dim Center(0 To 2) As Double
    Center(0) = CenterX:     Center(1) = CenterY:         Center(2) = CenterZ
    'Creando nuevo punto
    Set nombre = AutoCAD.Application.ActiveDocument.ModelSpace.AddPoint(Center)
    nombre.color = color
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

Public Function LineaPaperSpace(nombre As Variant, StartX As Variant, StartY As Variant, StartZ As Variant, EndX As Variant, EndY As Variant, EndZ As Variant, color As Variant, layer As String)
    'Declaración de Variables
    Dim PuntoA(0 To 2) As Double:    Dim PuntoB(0 To 2) As Double
    PuntoA(0) = StartX:     PuntoA(1) = StartY:         PuntoA(2) = StartZ
    PuntoB(0) = EndX:       PuntoB(1) = EndY:           PuntoB(2) = EndZ
    'Creando nueva linea
    Set nombre = AutoCAD.Application.ActiveDocument.PaperSpace.AddLine(PuntoA, PuntoB)
    nombre.color = color
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
    For I = 1 To df
    AutoCAD.Application.ActiveDocument.Close False
    Next I
End Function

Public Function Borrar(opcion As String)
    If opcion = "lineas" Then var = "AcDbLine"
    If opcion = "solidos" Then var = "AcDb3dSolid"
    If opcion = "puntos" Then var = "AcDbPoint"
    If opcion = "arcos" Then var = "AcDbArc"
    If opcion = "polilineas" Then var = "AcDbPolyline"
    If opcion = "bloques" Then var = "AcDbBlockReference"
    If opcion = "viewports" Then var = "AcDbViewport"
    If opcion = "circulos" Then var = "AcDbCircle"
    If opcion = "elipse" Then var = "AcDbEllipse"
    If opcion = "spline" Then var = "AcDbSpline"

    Dim Objeto As AcadEntity
    For Each Objeto In AutoCAD.Application.ActiveDocument.ModelSpace
        If Objeto.ObjectName = var Then
        Objeto.Delete
        End If
    Next Objeto
        AutoCAD.Application.ActiveDocument.SendCommand "regen" & vbCr
End Function

Public Function BorrarExepto(opcion As String)
    If opcion = "lineas" Then var = "AcDbLine"
    If opcion = "solidos" Then var = "AcDb3dSolid"
    If opcion = "puntos" Then var = "AcDbPoint"
    If opcion = "arcos" Then var = "AcDbArc"
    If opcion = "polilineas" Then var = "AcDbPolyline"
    If opcion = "bloques" Then var = "AcDbBlockReference"
    If opcion = "viewports" Then var = "AcDbViewport"
    If opcion = "circulos" Then var = "AcDbCircle"
    If opcion = "elipse" Then var = "AcDbEllipse"
    If opcion = "spline" Then var = "AcDbSpline"
    Dim Objeto As AcadEntity
    For Each Objeto In AutoCAD.Application.ActiveDocument.ModelSpace
        If Objeto.ObjectName = var Then
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
    I = 1
    Do While I > 0
    For Each Entidad In AutoCAD.Application.ActiveDocument.ModelSpace
    If Entidad.ObjectName = "AcDbBlockReference" Then
    Set bloques = Entidad
    bloques.Explode
    Entidad.Delete
    End If
    Next Entidad
    I = 0
    For Each Entidad In AutoCAD.Application.ActiveDocument.ModelSpace
    If Entidad.ObjectName = "AcDbBlockReference" Then
    I = 1
    End If
    Next Entidad
    For Each Entidad In AutoCAD.Application.ActiveDocument.ModelSpace
    If Entidad.ObjectName = "AcDbBlockReference" Then
    I = I + 1
    End If
    Next Entidad
    Loop
    AutoCAD.Application.ActiveDocument.PurgeAll 'Purgando
End Function

Public Function UCSWorld()
    AutoCAD.Application.ActiveDocument.SendCommand "ucs" & vbCr & "w" & vbCr 'UCS
End Function


Function activarAcad()



Dim objacadapp As AcadApplication

Set objacadapp = GetObject(, "Autocad.application")

AppActivate objacadapp.Caption

AutoCAD.WindowState = acMax
    
End Function
