VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EngCad_Autocad_Edition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim entidad As AcadEntity
Dim Block As AcadBlockReference
Public Sub BuscayExplotaBloques()
Dim entidad As AcadEntity
For Each entidad In Autocad.Application.ActiveDocument.modelspace
If entidad.ObjectName = "AcDbDbBlockReference" Then
Set Block = entidad
Bloc.Explode
End If
Next entidad
End Sub
Public Function ReemplazaPolilineasYTrazaCentros(Space As AcActiveSpace)
'esta macro reemplaza las polilineas por circulos, ademas traza lineas de eje
Dim entidad As AcadEntity
Dim Perimeter As Double
Dim Length As Variant
For Each entidad In Autocad.Application.ActiveDocument.modelspace
If entidad.ObjectName = "AcDbDbPolyline" Then
Length = Round(entidad.Length, 4)
End If
entidad.GetBoundingBox MinPoint, MaxPoint
Radio = ((MaxPoint(0) - MinPoint(0)) / 2)
Pi = 3.14159265358979
PerimetroPosible = Round(Radio * 2 * Pi, 4)
If PerimetroPosible = Length Then
d = Radio
Linea Space, MinPoint(0) - d, MinPoint(1) + d, MinPoint(2), MaxPoint(0) + d, MaxPoint(1) - d, MaxPoint(2), acByLayer, "0"
Linea Space, MinPoint(0) + d, MinPoint(1) - d, MinPoint(2), MaxPoint(0) - d, MaxPoint(1) + d, MaxPoint(2), acByLayer, "0"
Circulo Space, MinPoint(0) + d, MinPoint(1) + d, MinPoint(2), d, acByLayer, ".GENERAL ARRANGEMENT"
entidad.Delete
End If
Next entidad
ZoomExtents
End Function

Public Function BorrarEntidadesExepto(Ent As AcEntityName)
Dim ObjVBA As String
If Ent = 1 Then ObjVBA = "AcDb3dFace"
If Ent = 2 Then ObjVBA = "AcDb3dPolyline"
If Ent = 3 Then ObjVBA = "AcDb3dSolid"
If Ent = 4 Then ObjVBA = "AcDbArc"
If Ent = 5 Then ObjVBA = "AcDbAttribute"
If Ent = 6 Then ObjVBA = "AcDbAttributeReference"
If Ent = 7 Then ObjVBA = "AcDbBlockReference"
If Ent = 8 Then ObjVBA = "AcDbCircle"
If Ent = 47 Then ObjVBA = "AcDbDgnUnderlay"
If Ent = 41 Then ObjVBA = "AcDbDim3PointAngular"
If Ent = 9 Then ObjVBA = "AcDbDimAligned"
If Ent = 10 Then ObjVBA = "AcDbDimAngular"
If Ent = 44 Then ObjVBA = "AcDbDimArcLength"
If Ent = 12 Then ObjVBA = "AcDbDimDiametric"
If Ent = 13 Then ObjVBA = "AcDbDimOrdinate"
If Ent = 14 Then ObjVBA = "AcDbDimRadial"
If Ent = 45 Then ObjVBA = "AcDbDimRadialLarge"
If Ent = 15 Then ObjVBA = "AcDbDimRotated"
If Ent = 46 Then ObjVBA = "AcDbDwfUnderlay"
If Ent = 16 Then ObjVBA = "AcDbEllipse"
If Ent = 42 Then ObjVBA = "AcDbExternalReference"
If Ent = 17 Then ObjVBA = "AcDbHatch"
If Ent = 18 Then ObjVBA = "AcDbLeader"
If Ent = 19 Then ObjVBA = "AcDbLine"
If Ent = 38 Then ObjVBA = "AcDbMInsertBlock"
If Ent = 48 Then ObjVBA = "AcDbMLeader"
If Ent = 40 Then ObjVBA = "AcDbMLine"
If Ent = 21 Then ObjVBA = "AcDbMtext"
If Ent = 51 Then ObjVBA = "AcDbNurbSurface"
If Ent = 50 Then ObjVBA = "AcDbPdfUnderlay"
If Ent = 22 Then ObjVBA = "AcDbPoint"
If Ent = 39 Then ObjVBA = "AcDbPolyfaceMesh"
If Ent = 25 Then ObjVBA = "AcDbPolymesh"
If Ent = 23 Then ObjVBA = "AcDbPolyline"
If Ent = 35 Then ObjVBA = "AcDbPViewport"
If Ent = 26 Then ObjVBA = "AcDbRaster"
If Ent = 27 Then ObjVBA = "AcDbRay"
If Ent = 28 Then ObjVBA = "AcDbRegion"
If Ent = 29 Then ObjVBA = "AcDbShape"
If Ent = 30 Then ObjVBA = "AcDbSolid"
If Ent = 31 Then ObjVBA = "AcDbSpline"
If Ent = 49 Then ObjVBA = "AcDbSubDMesh"
If Ent = 43 Then ObjVBA = "AcDbTable"
If Ent = 32 Then ObjVBA = "AcDbText"
If Ent = 33 Then ObjVBA = "AcDbTolerance"
If Ent = 34 Then ObjVBA = "AcDbTrace"
If Ent = 36 Then ObjVBA = "AcDbXline"
Dim entidad As AcadEntity
For Each entidad In Autocad.Application.ActiveDocument.modelspace
If entidad.ObjectName = ObjVBA Then
Else
entidad.Delete
End If
Next entidad
Autocad.Application.ActiveDocument.SendCommand "regen" & vbCr
End Function

Public Function BorrarEntidades(Ent As AcEntityName)
Dim ObjVBA As String
If Ent = 1 Then ObjVBA = "AcDb3dFace"
If Ent = 2 Then ObjVBA = "AcDb3dPolyline"
If Ent = 3 Then ObjVBA = "AcDb3dSolid"
If Ent = 4 Then ObjVBA = "AcDbArc"
If Ent = 5 Then ObjVBA = "AcDbAttribute"
If Ent = 6 Then ObjVBA = "AcDbAttributeReference"
If Ent = 7 Then ObjVBA = "AcDbBlockReference"
If Ent = 8 Then ObjVBA = "AcDbCircle"
If Ent = 47 Then ObjVBA = "AcDbDgnUnderlay"
If Ent = 41 Then ObjVBA = "AcDbDim3PointAngular"
If Ent = 9 Then ObjVBA = "AcDbDimAligned"
If Ent = 10 Then ObjVBA = "AcDbDimAngular"
If Ent = 44 Then ObjVBA = "AcDbDimArcLength"
If Ent = 12 Then ObjVBA = "AcDbDimDiametric"
If Ent = 13 Then ObjVBA = "AcDbDimOrdinate"
If Ent = 14 Then ObjVBA = "AcDbDimRadial"
If Ent = 45 Then ObjVBA = "AcDbDimRadialLarge"
If Ent = 15 Then ObjVBA = "AcDbDimRotated"
If Ent = 46 Then ObjVBA = "AcDbDwfUnderlay"
If Ent = 16 Then ObjVBA = "AcDbEllipse"
If Ent = 42 Then ObjVBA = "AcDbExternalReference"
If Ent = 17 Then ObjVBA = "AcDbHatch"
If Ent = 18 Then ObjVBA = "AcDbLeader"
If Ent = 19 Then ObjVBA = "AcDbLine"
If Ent = 38 Then ObjVBA = "AcDbMInsertBlock"
If Ent = 48 Then ObjVBA = "AcDbMLeader"
If Ent = 40 Then ObjVBA = "AcDbMLine"
If Ent = 21 Then ObjVBA = "AcDbMtext"
If Ent = 51 Then ObjVBA = "AcDbNurbSurface"
If Ent = 50 Then ObjVBA = "AcDbPdfUnderlay"
If Ent = 22 Then ObjVBA = "AcDbPoint"
If Ent = 39 Then ObjVBA = "AcDbPolyfaceMesh"
If Ent = 25 Then ObjVBA = "AcDbPolymesh"
If Ent = 23 Then ObjVBA = "AcDbPolyline"
If Ent = 35 Then ObjVBA = "AcDbPViewport"
If Ent = 26 Then ObjVBA = "AcDbRaster"
If Ent = 27 Then ObjVBA = "AcDbRay"
If Ent = 28 Then ObjVBA = "AcDbRegion"
If Ent = 29 Then ObjVBA = "AcDbShape"
If Ent = 30 Then ObjVBA = "AcDbSolid"
If Ent = 31 Then ObjVBA = "AcDbSpline"
If Ent = 49 Then ObjVBA = "AcDbSubDMesh"
If Ent = 43 Then ObjVBA = "AcDbTable"
If Ent = 32 Then ObjVBA = "AcDbText"
If Ent = 33 Then ObjVBA = "AcDbTolerance"
If Ent = 34 Then ObjVBA = "AcDbTrace"
If Ent = 36 Then ObjVBA = "AcDbXline"
Dim entidad As AcadEntity
For Each entidad In Autocad.Application.ActiveDocument.modelspace
If entidad.ObjectName = ObjVBA Then
entidad.Delete
Else
End If
Next entidad
Autocad.Application.ActiveDocument.SendCommand "regen" & vbCr
End Function
Public Sub EliminarTodo()
e = Autocad.Application.ActiveDocument.modelspace.Count
If e > 0 Then
Autocad.Application.ActiveDocument.SendCommand "_ai_selall" & vbCr & "_.erase" & vbCr
Else:
End If
End Sub
Public Sub NuevoDibujo()
Autocad.Application.Documents.Add (acad)
End Sub
Public Function Seleccion(Name As Variant, StartX As Variant, StartY As Variant, StartZ As Variant, EndX As Variant, EndY As Variant, EndZ As Variant)
Dim ssetObj As AcadSelectionSet
Dim mode As Integer: Dim PuntoA(0 To 2) As Double: Dim PuntoB(0 To 2) As Double
Set ssetObj = Autocad.Application.ActiveDocument.SelectionSets.Add(Name)
mode = acSelectionSetCrossing
PuntoA(0) = StartX:     PuntoA(1) = StartY:         PuntoA(2) = StartZ
PuntoB(0) = EndX:       PuntoB(1) = EndY:           PuntoB(2) = EndZ
ssetObj.Select mode, PuntoA, PuntoB
End Function

Public Function CrearBloque(Name As Variant, StartX As Variant, StartY As Variant, StartZ As Variant)
Dim blockObj As AcadBlock: Dim PuntoA(0 To 2) As Double
PuntoA(0) = StartX: PuntoA(1) = StartY: PuntoA(2) = StartZ
Set blockObj = Autocad.ActiveDocument.Blocks.Add(PuntoA, Name)
End Function


Public Function InsertBlockName(Space As AcActiveSpace, Name As String, StartX As Variant, StartY As Variant, StartZ As Variant)
Dim blockRefObj As AcadBlockReference: Dim PuntoA(0 To 2) As Double
PuntoA(0) = StartX: PuntoA(1) = StartY: PuntoA(2) = StartZ
If Space = 0 Then Set blockRefObj = Autocad.Application.ActiveDocument.PaperSpace.InsertBlock(PuntoA, Name, 1#, 1#, 1#, 0)
If Space = 1 Then Set blockRefObj = Autocad.Application.ActiveDocument.modelspace.InsertBlock(PuntoA, Name, 1#, 1#, 1#, 0)
End Function






