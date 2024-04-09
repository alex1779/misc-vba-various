VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EngCad_Autocad_General"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Enum Espacios
modelspace
PaperSpace
End Enum

Public Enum UCSTypes
Previous
view
World
End Enum

Public Enum Views
Top
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
Shaded
shaded_with_Edges
shades_of_Gray
SKetchy
X_ray
End Enum


Public Function AbrirArchivo(PathFile As Variant)
Autocad.Application.Documents.Open (PathFile)
End Function
Public Function ActivarSpace(Space As AcActiveSpace)
Autocad.Application.ActiveDocument.ActiveSpace = Space
End Function
Public Function ContarEntidades(Espacio As Espacios)
If Espacio = 0 Then ContarEntidades = Autocad.Application.ActiveDocument.modelspace.Count
If Espacio = 1 Then ContarEntidades = Autocad.Application.ActiveDocument.PaperSpace.Count
End Function
Public Function GuardarComo(Path As Variant, Filename As Variant, SaveType As AcSaveAsType)
Autocad.Application.ActiveDocument.SaveAs (Path & Filename), SaveType
End Function
Public Function ucs(UcsType As UCSTypes)
If UcsType = 0 Then UcsT = "P"
If UcsType = 1 Then UcsT = "V"
If UcsType = 2 Then UcsT = "W"
Autocad.Application.ActiveDocument.SendCommand "_UCS" & vbCr & UcsT & vbCr
End Function
Public Sub BorrarSelectionSets()
On Error Resume Next
Dim ssetObj As AcadSelectionSet
For Each ssetObj In Autocad.Application.ActiveDocument.SelectionSets
ssetObj.Delete
Next
End Sub
Public Sub Cerrar()
Autocad.Application.ActiveDocument.Close
End Sub
Public Sub GridOff()
Autocad.Application.ActiveDocument.SendCommand "grid" & vbCr & "of" & vbCr
End Sub
Public Sub Guardar()
Autocad.Application.ActiveDocument.Save
End Sub
Public Sub JointAll()
Autocad.Application.ActiveDocument.SendCommand "_ai_selall" & vbCr & "_j" & vbCr
End Sub
Public Sub NombreDeEntidad()
Dim entidad As AcadEntity
Set ssetObj = Autocad.Application.ActiveDocument.SelectionSets.Add(Rnd)
ssetObj.SelectOnScreen
For Each entidad In ssetObj
MsgBox entidad.ObjectName
Next
ssetObj.Delete
End Sub
Public Sub Overkill()
Autocad.Application.ActiveDocument.SendCommand "-overkill" & vbCr & "all" & vbCr & vbCr & vbCr
End Sub
Public Sub Purge()
Autocad.Application.ActiveDocument.PurgeAll
End Sub
Public Sub Regen()
Autocad.Application.ActiveDocument.SendCommand "regen" & vbCr
End Sub
Public Sub SelectAll()
Autocad.Application.ActiveDocument.SendCommand "_ai_selall" & vbCr
End Sub
Public Sub UnionAllSolids()
Autocad.Application.ActiveDocument.SendCommand "_ai_selall" & vbCr & "_union" & vbCr & "all" & vbCr & vbCr
End Sub
Public Sub CopyAllFromZero()
Autocad.Application.ActiveDocument.SendCommand "_ai_selall" & vbCr & "_copybase" & vbCr & "0,0,0" & vbCr
End Sub
Public Sub PasteAllAtZero()
Autocad.Application.ActiveDocument.SendCommand "_pasteclip " & vbCr & "0,0,0 " & vbCr
End Sub
Public Function Pegar(Punto As String)
Autocad.Application.ActiveDocument.SendCommand "_pasteclip " & vbCr & Punto & vbCr
End Function
Public Function Vista(Ent As Views)
If Ent = 0 Then ObjVBA = "_Top"
If Ent = 1 Then ObjVBA = "_Bottom"
If Ent = 2 Then ObjVBA = "_Left"
If Ent = 3 Then ObjVBA = "_Right"
If Ent = 4 Then ObjVBA = "_Front"
If Ent = 5 Then ObjVBA = "_Back"
If Ent = 6 Then ObjVBA = "_swiso"
If Ent = 7 Then ObjVBA = "_seiso"
If Ent = 8 Then ObjVBA = "_neiso"
If Ent = 9 Then ObjVBA = "_nwiso"
Autocad.Application.ActiveDocument.SendCommand "_-view" & vbCr & ObjVBA & vbCr
End Function
Public Function Sombreado(Ent As Sombras)
If Ent = 0 Then ObjVBA = "2"
If Ent = 1 Then ObjVBA = "W"
If Ent = 2 Then ObjVBA = "H"
If Ent = 3 Then ObjVBA = "_R"
If Ent = 4 Then ObjVBA = "_C"
If Ent = 5 Then ObjVBA = "S"
If Ent = 6 Then ObjVBA = "G"
If Ent = 7 Then ObjVBA = "SK"
If Ent = 8 Then ObjVBA = "X"
Autocad.Application.ActiveDocument.SendCommand "shademode" & vbCr & ObjVBA & vbCr
End Function
Public Function Random(menor As Integer, mayor As Integer)
Random = ((mayor - menor + 1) * Rnd + menor)
Do While Random = 0: Random = ((mayor - menor + 1) * Rnd + menor): Loop 'Non Zero
Do While Random > mayor: Random = ((mayor - menor + 1) * Rnd + menor): Loop 'Non Zero
End Function
