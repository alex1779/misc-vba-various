Attribute VB_Name = "Generalis"
Sub General()


    Dim Spline As AcadSpline
    Dim LayerName As String
    
    Set modelspace = AutoCAD.Application.ActiveDocument.modelspace
    i = 1
    
    'Eliminando Duplicados
    Call Overkill
    'Cambiando a Metodo Fit
    On Error Resume Next
    For Each Spline In modelspace
    NuevoLayer LayerNuevo, Spline.ObjectID, 256, "Continuous", 0, True
    Spline.SplineMethod = acFit
    Spline.layer = Spline.ObjectID
    Spline.Color = acByLayer
    Next
    'Purgando
    AutoCAD.Application.ActiveDocument.PurgeAll
    
    

    For Each Spline In modelspace

'    LayerName = Spline.ObjectID
'    'Apaga todos los layers excepto...
'        For Each Spline In modelspace

    ApagaTodosLosLayersExcepto Spline.ObjectID
    ZoomExtents

    AutoCAD.Application.ActiveDocument.SendCommand "_ai_selall" & vbCr  'Select All

    AutoCAD.Application.ActiveDocument.SendCommand "SPLINEDIT" & vbCr & "p" & vbCr & "p" & vbCr & 10 & vbCr
    


    i = i + 1
    Next




    
    
End Sub

Sub Pruebaaaaa()

    NuevoLayer LayerNuevo, "Alejandro", 1, "Continuous", 0, True
    Activalayer "Alejandro"



End Sub


Public Function ApagaTodosLosLayersExcepto(Name As String)
    On Error Resume Next
    Dim layerObj As AcadLayer
    For Each layerObj In AutoCAD.Application.ActiveDocument.Layers
    If layerObj.Name = Name Then
    Else:
    layerObj.Freeze = True
    layerObj.LayerOn = False
    End If
    Next layerObj
    AutoCAD.Application.ActiveDocument.SendCommand "regen" & vbCr 'Regen
End Function
