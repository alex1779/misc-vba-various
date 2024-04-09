VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EngCad_Autocad_Draw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function Arc(Space As AcActiveSpace, CenterX As Variant, CenterY As Variant, CenterZ As Variant, Radious As Variant, StartAngle As Variant, EndAngle As Variant, Optional Color As AcColor, Optional layer As String)
Dim ArcObj As AcadArc
Dim center(0 To 2) As Double
StartAngle = (StartAngle * 3.14159265358979) / 180
EndAngle = (EndAngle * 3.14159265358979) / 180
center(0) = CenterX: center(1) = CenterY: center(2) = CenterZ
If Space = 0 Then Set ArcObj = Autocad.Application.ActiveDocument.PaperSpace.AddArc(center, Radious, StartAngle, EndAngle)
If Space = 1 Then Set ArcObj = Autocad.Application.ActiveDocument.modelspace.AddArc(center, Radious, StartAngle, EndAngle)
On Error Resume Next
ArcObj.Color = Color
ArcObj.layer = layer
End Function

Public Function Linea(Space As AcActiveSpace, StartX As Variant, StartY As Variant, StartZ As Variant, EndX As Variant, EndY As Variant, EndZ As Variant, Optional Color As AcColor, Optional layer As String)
Dim LineObj As AcadLine: Dim StartPoint(0 To 2) As Double: Dim EndPoint(0 To 2) As Double
StartPoint(0) = StartX: StartPoint(1) = StartY: StartPoint(2) = StartZ
EndPoint(0) = EndX: EndPoint(1) = EndY: EndPoint(2) = EndZ
If Space = 0 Then Set LineObj = Autocad.Application.ActiveDocument.PaperSpace.AddLine(StartPoint, EndPoint)
If Space = 1 Then Set LineObj = Autocad.Application.ActiveDocument.modelspace.AddLine(StartPoint, EndPoint)
On Error Resume Next
LineObj.Color = Color
LineObj.layer = layer
End Function

Public Function Point(Space As AcActiveSpace, CenterX As Variant, CenterY As Variant, CenterZ As Variant, Optional Color As AcColor, Optional layer As String)
Dim PointObj As AcadPoint
Dim Pnt(0 To 2) As Double
Pnt(0) = CenterX: Pnt(1) = CenterY: Pnt(2) = CenterZ
If Space = 0 Then Set PointObj = Autocad.Application.ActiveDocument.PaperSpace.AddPoint(Pnt)
If Space = 1 Then Set PointObj = Autocad.Application.ActiveDocument.modelspace.AddPoint(Pnt)
On Error Resume Next
PointObj.Color = Color
PointObj.layer = layer
End Function

Public Function Circulo(Space As AcActiveSpace, CenterX As Variant, CenterY As Variant, CenterZ As Variant, Radious As Variant, Optional Color As AcColor, Optional layer As String)
Dim CirculoObj As AcadCircle
Dim Pnt(0 To 2) As Double
Pnt(0) = CenterX: Pnt(1) = CenterY: Pnt(2) = CenterZ
If Space = 0 Then Set CirculoObj = Autocad.Application.ActiveDocument.PaperSpace.AddCircle(Pnt, Radious)
If Space = 1 Then Set CirculoObj = Autocad.Application.ActiveDocument.modelspace.AddCircle(Pnt, Radious)
On Error Resume Next
CirculoObj.Color = Color
CirculoObj.layer = layer
End Function

Public Function XLine(Space As AcActiveSpace, StartX As Variant, StartY As Variant, StartZ As Variant, EndX As Variant, EndY As Variant, EndZ As Double, Optional Color As AcColor, Optional layer As String)
Dim XLineObj As AcadXline: Dim StartPoint(0 To 2) As Double: Dim EndPoint(0 To 2) As Double
StartPoint(0) = StartX: StartPoint(1) = StartY: StartPoint(2) = StartZ
EndPoint(0) = EndX: EndPoint(1) = EndY: EndPoint(2) = EndZ
If Space = 0 Then Set XLineObj = Autocad.Application.ActiveDocument.PaperSpace.AddXline(StartPoint, EndPoint)
If Space = 1 Then Set XLineObj = Autocad.Application.ActiveDocument.modelspace.AddXline(StartPoint, EndPoint)
On Error Resume Next
XLineObj.Color = Color
XLineObj.layer = layer
End Function
Public Function MText(Space As AcActiveSpace, StartX As Variant, StartY As Variant, StartZ As Variant, Width As Variant, CadenaTexto As Variant, Optional Color As AcColor, Optional layer As String)
Dim MTextObj As AcadMText: Dim PuntoA(0 To 2) As Double
PuntoA(0) = StartX: PuntoA(1) = StartY: PuntoA(2) = StartZ
If Space = 0 Then Set MTextObj = Autocad.Application.ActiveDocument.PaperSpace.AddMText(PuntoA, Width, CadenaTexto)
If Space = 1 Then Set MTextObj = Autocad.Application.ActiveDocument.modelspace.AddMText(PuntoA, Width, CadenaTexto)
On Error Resume Next
With MTextObj
.AttachmentPoint = acAttachmentPointMiddleLeft
.Color = Color
.insertionPoint = PuntoA
End With
On Error Resume Next
MTextObj.Color = Color
MTextObj.layer = layer
End Function
Public Function Rectangulo(Space As AcActiveSpace, Ancho As Variant, Alto As Variant, BasePointX As Variant, BasePointY As Variant, Optional Color As AcColor, Optional layer As String)
Linea Space, BasePointX, BasePointY, 0, BasePointX + Ancho, BasePointY, 0
Linea Space, BasePointX, BasePointY + Alto, 0, BasePointX + Ancho, BasePointY + Alto, 0
Linea Space, BasePointX, BasePointY, 0, BasePointX, BasePointY + Alto, 0
Linea Space, BasePointX + Ancho, BasePointY, 0, BasePointX + Ancho, BasePointY + Alto, 0
Autocad.Application.ActiveDocument.SendCommand "regen" & vbCr
End Function
Public Function RectanguloPolyClosed(Space As AcActiveSpace, Ancho As Variant, Alto As Variant, BasePointX As Variant, BasePointY As Variant, Optional Color As AcColor, Optional layer As String)
Dim Poly As AcadPolyline
Dim PuntoA(0 To 11) As Double
PuntoA(0) = BasePointX:             PuntoA(1) = BasePointY:             PuntoA(2) = 0
PuntoA(3) = BasePointX:             PuntoA(4) = BasePointY + Alto:      PuntoA(5) = 0
PuntoA(6) = BasePointX + Ancho:     PuntoA(7) = BasePointY + Alto:      PuntoA(8) = 0
PuntoA(9) = BasePointX + Ancho:     PuntoA(10) = BasePointY:      PuntoA(11) = 0
Set Poly = Autocad.Application.ActiveDocument.modelspace.AddPolyline(PuntoA)
Poly.Closed = True
On Error Resume Next
Poly.Color = Color
Poly.layer = layer
Autocad.Application.ActiveDocument.SendCommand "regen" & vbCr
End Function
Public Function Face3d(Space As AcActiveSpace, Ancho As Variant, Alto As Variant, BasePointX As Variant, BasePointY As Variant, Optional Color As AcColor, Optional layer As String)
Dim Face3dObj As Acad3DFace: Dim PuntoA(0 To 2) As Double: Dim PuntoB(0 To 2) As Double: Dim PuntoC(0 To 2) As Double: Dim PuntoD(0 To 2) As Double
PuntoA(0) = BasePointX:             PuntoA(1) = BasePointY:             PuntoA(2) = 0
PuntoB(0) = BasePointX:             PuntoB(1) = BasePointY + Alto:      PuntoB(2) = 0
PuntoC(0) = BasePointX + Ancho:     PuntoC(1) = BasePointY + Alto:      PuntoC(2) = 0
PuntoD(0) = BasePointX + Ancho:     PuntoD(1) = BasePointY:             PuntoD(2) = 0
If Space = 0 Then Set Face3dObj = Autocad.Application.ActiveDocument.PaperSpace.Add3DFace(PuntoA, PuntoB, PuntoC, PuntoD)
If Space = 1 Then Set Face3dObj = Autocad.Application.ActiveDocument.modelspace.Add3DFace(PuntoA, PuntoB, PuntoC, PuntoD)
On Error Resume Next
Face3dObj.Color = Color
Face3dObj.layer = layer
Autocad.Application.ActiveDocument.SendCommand "regen" & vbCr
End Function

Public Function Cilindro3d(Space As AcActiveSpace, CenterX As Variant, CenterY As Variant, CenterZ As Variant, Radious As Variant, Altura As Double, Direccion As Double, Optional Color As AcColor, Optional layer As String)
Dim Cilindro3dObj As Acad3DSolid: Dim Centro(0 To 2) As Double: Dim regionObj As Variant: Dim Circulo(0) As AcadCircle
Centro(0) = CenterX: Centro(1) = CenterY: Centro(2) = CenterZ
If Space = 0 Then Set Circulo(0) = Autocad.Application.ActiveDocument.PaperSpace.AddCircle(Centro, Radious)
If Space = 1 Then Set Circulo(0) = Autocad.Application.ActiveDocument.modelspace.AddCircle(Centro, Radious)
If Space = 0 Then regionObj = Autocad.Application.ActiveDocument.PaperSpace.AddRegion(Circulo)
If Space = 1 Then regionObj = Autocad.Application.ActiveDocument.modelspace.AddRegion(Circulo)
If Space = 0 Then Set Cilindro3dObj = Autocad.Application.ActiveDocument.PaperSpace.AddExtrudedSolid(regionObj(0), Altura, Direccion)
If Space = 1 Then Set Cilindro3dObj = Autocad.Application.ActiveDocument.modelspace.AddExtrudedSolid(regionObj(0), Altura, Direccion)
Circulo(0).Delete
regionObj(0).Delete
On Error Resume Next
Cilindro3dObj.Color = Color
Cilindro3dObj.layer = layer
Autocad.Application.ActiveDocument.SendCommand "regen" & vbCr
End Function

