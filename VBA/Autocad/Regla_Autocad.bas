Attribute VB_Name = "Regla_Autocad"
Option Explicit

Public Function Regla(origenx As Double, origeny As Double, ancho As Integer, alto As Integer)
    On Error Resume Next
'    AutoCAD.Application.ActiveDocument.SendCommand "_ai_selall" & vbCr & "_.erase" & vbCr 'Select All & Erase
'    Dim ancho As Integer
'    Dim alto As Integer
    Dim linea1 As Variant
    Dim i As Integer
    Dim PuntoA(0 To 2) As Double: Dim mtextObj As AcadMText: Dim Width As Double: Dim CadenaTexto As String 'Declaración de Variables
'    ancho = 100
'    alto = 100
'    ancho = AutoCAD.Application.ActiveDocument.Utility.GetReal("Ingrese Ancho de la regla: ")
'    alto = AutoCAD.Application.ActiveDocument.Utility.GetReal("Ingrese Alto de la regla: ")
    Linea linea1, origenx, origeny, 0, origenx + ancho, origeny, 0, 7, "0"
    Linea linea1, origenx, origeny, 0, origenx, origeny + alto, 0, 7, "0"


    For i = 0 To ancho / 10
    Linea linea1, origenx + i * 10, origeny + 0, 0, origenx + i * 10, origeny + -5, 0, 7, "0"
    PuntoA(0) = origenx + i * 10: PuntoA(1) = origeny + -7: PuntoA(2) = 0     'Iniciación de Variables
    Width = 4
    CadenaTexto = i
    Set mtextObj = AutoCAD.Application.ActiveDocument.ModelSpace.AddMText(PuntoA, Width, CadenaTexto)    'Creando Nuevo Texto
    With mtextObj
        .AttachmentPoint = acAttachmentPointMiddleCenter
        .Color = 7
        .InsertionPoint = PuntoA
    End With
    Next i


    For i = 0 To alto / 10
    Linea linea1, origenx + 0, origeny + i * 10, 0, origenx + -5, origeny + i * 10, 0, 7, "0"
    PuntoA(0) = origenx + -7: PuntoA(1) = origeny + i * 10: PuntoA(2) = 0     'Iniciación de Variables
    Width = 4
    CadenaTexto = i
    Set mtextObj = AutoCAD.Application.ActiveDocument.ModelSpace.AddMText(PuntoA, Width, CadenaTexto)    'Creando Nuevo Texto
    With mtextObj
        .AttachmentPoint = acAttachmentPointMiddleCenter
        .Color = 7
        .InsertionPoint = PuntoA
    End With
    Next i



    For i = 0 To ancho
    Linea linea1, origenx + i, origeny + 0, 0, origenx + i, origeny + -2, 0, 7, "0"
    Next i
    For i = 0 To alto
    Linea linea1, origenx + 0, origeny + i, 0, origenx + -2, origeny + i, 0, 7, "0"
    Next i
    
    ZoomExtents    'Zoom Extents
    
End Function

Private Function Linea(nombre As Variant, StartX As Variant, StartY As Variant, StartZ As Variant, EndX As Variant, EndY As Variant, EndZ As Double, Color As Variant, layer As String)
    'Declaración de Variables
    Dim PuntoA(0 To 2) As Double:    Dim PuntoB(0 To 2) As Double
    PuntoA(0) = StartX:     PuntoA(1) = StartY:         PuntoA(2) = StartZ
    PuntoB(0) = EndX:       PuntoB(1) = EndY:           PuntoB(2) = EndZ
    'Creando nueva linea
    Set nombre = AutoCAD.Application.ActiveDocument.ModelSpace.AddLine(PuntoA, PuntoB)
    nombre.Color = Color
    nombre.layer = layer
End Function



