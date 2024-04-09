Attribute VB_Name = "FigurasGeometricas"
Public Function Linea(nombre As Variant, StartX As Variant, StartY As Variant, StartZ As Variant, EndX As Variant, EndY As Variant, EndZ As Variant, Color As Variant, layer As String)
    'Declaración de Variables
    Dim PuntoA(0 To 2) As Double:    Dim PuntoB(0 To 2) As Double
    PuntoA(0) = StartX:     PuntoA(1) = StartY:         PuntoA(2) = StartZ
    PuntoB(0) = EndX:       PuntoB(1) = EndY:           PuntoB(2) = EndZ
    'Creando nueva linea
    Set nombre = AutoCAD.Application.ActiveDocument.ModelSpace.AddLine(PuntoA, PuntoB)
    nombre.Color = Color
    nombre.layer = layer
End Function
Public Function Circulo(nombre As Variant, CenterX As Variant, CenterY As Variant, CenterZ As Variant, Radious As Variant, Color As Variant, layer As String)
    'Declaración de Variables
    ''Dim Circulo As AcadCircle
    Dim Center(0 To 2) As Double
    Center(0) = CenterX:     Center(1) = CenterY:         Center(2) = CenterZ

    'Creando nuevo círculo
    Set nombre = AutoCAD.Application.ActiveDocument.ModelSpace.AddCircle(Center, Radious)
    nombre.Color = Color
    nombre.layer = layer
End Function

Public Function Arco(nombre As Variant, CentroX As Variant, CentroY As Variant, CentroZ As Variant, Radious As Variant, StartAngle As Variant, EndAngle As Variant, Color As Variant, layer As String)
    'Declaración de Variables
''    Dim Arco As AcadArc:
    Dim CentroArcoA(0 To 2) As Double
    Radious = Radious
    StartAngle = StartAngle
    EndAngle = EndAngle
    CentroArcoA(0) = CentroX: CentroArcoA(1) = CentroY: CentroArcoA(2) = CentroZ
    'Dibujando arco
    Set nombre = AutoCAD.Application.ActiveDocument.ModelSpace.AddArc(CentroArcoA, Radious, StartAngle, EndAngle)
    nombre.Color = Color
    nombre.layer = layer
End Function

Public Function Punto(nombre As Variant, CenterX As Variant, CenterY As Variant, CenterZ As Variant, Color As Variant, layer As String)
    'Declaración de Variables
    Dim Center(0 To 2) As Double
    Center(0) = CenterX:     Center(1) = CenterY:         Center(2) = CenterZ
    'Creando nuevo punto
    Set nombre = AutoCAD.Application.ActiveDocument.ModelSpace.AddPoint(Center)
    nombre.Color = Color
    nombre.layer = layer
End Function
Public Function XLinea(nombre As Variant, StartX As Variant, StartY As Variant, StartZ As Variant, EndX As Variant, EndY As Variant, EndZ As Double, Color As Variant, layer As String)
    'Declaración de Variables
    Dim PuntoA(0 To 2) As Double:    Dim PuntoB(0 To 2) As Double
    PuntoA(0) = StartX:     PuntoA(1) = StartY:         PuntoA(2) = StartZ
    PuntoB(0) = EndX:       PuntoB(1) = EndY:           PuntoB(2) = EndZ
    'Creando nueva linea
    Set nombre = AutoCAD.Application.ActiveDocument.ModelSpace.AddXline(PuntoA, PuntoB)
    nombre.Color = Color
    nombre.layer = layer
End Function



