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
Set layerObj = Autocad.Application.ActiveDocument.Layers.Item(Name)
Autocad.Application.ActiveDocument.ActiveLayer = layerObj
End Function
Public Function Apagar(Name As Variant)
Set layerObj = Autocad.Application.ActiveDocument.Layers.Item(nombre)
layerObj.Freeze = True
End Function
Public Function Encender(Name As Variant)
Set layerObj = Autocad.Application.ActiveDocument.Layers.Item(nombre)
layerObj.Freeze = False
End Function
Public Function ApagarTodos()
On Error Resume Next
Dim layerObj As AcadLayer
For Each layerObj In Autocad.Application.ActiveDocument.Layers
layerObj.Freeze = True
layerObj.LayerOn = False
Next layerObj
End Function
Public Function EncenderTodos()
On Error Resume Next
Dim layerObj As AcadLayer
For Each layerObj In Autocad.Application.ActiveDocument.Layers
layerObj.Freeze = False
layerObj.LayerOn = True
Next layerObj
End Function
Public Function Nuevo(Name As Variant)
Dim NewLayer As AcadLayer
Set NewLayer = Autocad.Application.ActiveDocument.Layers.Add(Name)
End Function
Public Function NuevoCompleto(Name As String, Color As AcColor, LineWeight As AcLineWeight, Plottable As Boolean, Optional linetype As String)
Dim NewLayer As AcadLayer
Set NewLayer = Autocad.Application.ActiveDocument.Layers.Add(Name)
NewLayer.Color = Color
NewLayer.LineWeight = LineWeight
NewLayer.Plottable = Plottable
NewLayer.Name = Name
On Error Resume Next
NewLayer.linetype = linetype
End Function




