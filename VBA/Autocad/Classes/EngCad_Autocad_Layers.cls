VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EngCad_Autocad_Layers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim layerObj As AcadLayer
Public Function Activar(Name As Variant)
Set layerObj = AutoCAD.Application.ActiveDocument.Layers.Item(Name)
AutoCAD.Application.ActiveDocument.ActiveLayer = layerObj
AutoCAD.Application.ActiveDocument.SendCommand "regen" & vbCr
End Function
Public Function Apagar(Name As Variant)
Set layerObj = AutoCAD.Application.ActiveDocument.Layers.Item(nombre)
layerObj.Freeze = True
AutoCAD.Application.ActiveDocument.SendCommand "regen" & vbCr
End Function
Public Function Encender(Name As Variant)
Set layerObj = AutoCAD.Application.ActiveDocument.Layers.Item(nombre)
layerObj.Freeze = False
AutoCAD.Application.ActiveDocument.SendCommand "regen" & vbCr
End Function
Public Function Eliminar(Name As Variant)
Set layerObj = AutoCAD.Application.ActiveDocument.Layers.Item(nombre)
layerObj.Delete
AutoCAD.Application.ActiveDocument.SendCommand "regen" & vbCr
End Function
Public Function ApagarTodos()
On Error Resume Next
Dim layerObj As AcadLayer
For Each layerObj In AutoCAD.Application.ActiveDocument.Layers
layerObj.Freeze = True
layerObj.LayerOn = False
Next layerObj
AutoCAD.Application.ActiveDocument.SendCommand "regen" & vbCr
End Function
Public Function EncenderTodos()
On Error Resume Next
Dim layerObj As AcadLayer
For Each layerObj In AutoCAD.Application.ActiveDocument.Layers
layerObj.Freeze = False
layerObj.LayerOn = True
Next layerObj
AutoCAD.Application.ActiveDocument.SendCommand "regen" & vbCr
End Function
Public Function layer(nombre As Variant, Name As String, Color As Variant, LineType As String, LineWeight As Variant, Plottable As Boolean)
    Set nombre = AutoCAD.Application.ActiveDocument.Layers.Add(Name)
    nombre.Color = Color
    nombre.LineType = LineType
    nombre.LineWeight = LineWeight
    nombre.Plottable = Plottable
    nombre = nombre.Name
End Function

Public Function Nuevo(Name As Variant)
Dim NewLayer As AcadLayer
Set NewLayer = AutoCAD.Application.ActiveDocument.Layers.Add(Name)
AutoCAD.Application.ActiveDocument.SendCommand "regen" & vbCr
End Function
Public Function NuevoCompleto(Name As String, Color As AcColor, LineWeight As AcLineWeight, Plottable As Boolean, Optional LineType As String)
Dim NewLayer As AcadLayer
Set NewLayer = AutoCAD.Application.ActiveDocument.Layers.Add(Name)
NewLayer.Color = Color
NewLayer.LineWeight = LineWeight
NewLayer.Plottable = Plottable
NewLayer.Name = Name
On Error Resume Next
NewLayer.LineType = LineType
End Function
Public Function ApagaExcepto(nombre As Variant)
Dim layerObj As AcadLayer
For Each layerObj In AutoCAD.Application.ActiveDocument.Layers
    If layerObj.Name <> nombre Then
    layerObj.Freeze = True
    layerObj.LayerOn = False
    Else
    AutoCAD.Application.ActiveDocument.SendCommand "regen" & vbCr
    End If
Next
End Function
Public Sub CargarLayersAlliedPumps()
On Error Resume Next
    AutoCAD.Application.ActiveDocument.Linetypes.Load "CENTER", "acad.lin"
    AutoCAD.Application.ActiveDocument.Linetypes.Load "HIDDEN", "acad.lin"
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
Public Sub CargarLayersEntidades2d()
' CARGA UN LAYER PARA CADA ENTIDAD 2D. ESTABLECE COLORES DEL 1 AL 20.
On Error Resume Next
'Entidades 2d
layer LayerObj19, ".Line", 1, "Continuous", acLnWtByLwDefault, True
layer LayerObj23, ".Polyline", 2, "Continuous", acLnWtByLwDefault, True
layer LayerObj23, ".Polyline2d", 3, "Continuous", acLnWtByLwDefault, True
layer LayerObj2, ".Polyline3d", 4, "Continuous", acLnWtByLwDefault, True
layer LayerObj4, ".Arc", 5, "Continuous", acLnWtByLwDefault, True
layer LayerObj8, ".Circle", 6, "Continuous", acLnWtByLwDefault, True
layer LayerObj31, ".Spline", 8, "Continuous", acLnWtByLwDefault, True
layer LayerObj16, ".Ellipse", 9, "Continuous", acLnWtByLwDefault, True
layer LayerObj36, ".Xline", 10, "Continuous", acLnWtByLwDefault, True
layer LayerObj28, ".Region", 11, "Continuous", acLnWtByLwDefault, True
layer LayerObj22, ".Point", 12, "Continuous", acLnWtByLwDefault, True
layer LayerObj17, ".Hatch", 13, "Continuous", acLnWtByLwDefault, True
layer LayerObj21, ".Mtext", 14, "Continuous", acLnWtByLwDefault, True
layer LayerObj32, ".Text", 15, "Continuous", acLnWtByLwDefault, True
layer LayerObj18, ".Leader", 16, "Continuous", acLnWtByLwDefault, True
layer LayerObj48, ".MLeader", 17, "Continuous", acLnWtByLwDefault, True
layer LayerObj40, ".MLine", 18, "Continuous", acLnWtByLwDefault, True
layer LayerObj26, ".Raster", 19, "Continuous", acLnWtByLwDefault, True
layer LayerObj27, ".Ray", 20, "Continuous", acLnWtByLwDefault, True
End Sub
