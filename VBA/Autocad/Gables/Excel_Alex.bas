Attribute VB_Name = "Excel_Alex"
Public Function NuevaHoja(nombre As String)
    Application.DisplayAlerts = False
    
    For Each Worksheet In Worksheets
        If Worksheet.Name = nombre Then
        Sheets(nombre).Select
        Titulos
        Exit Function
        Else:
        existe = False
        End If
    Next
    
    If existe = False Then
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = nombre
    Titulos
    Else:
    End If
    
    Application.DisplayAlerts = True
    



End Function

Public Function BorrarHojasVacias()
Application.DisplayAlerts = False

For Each Worksheet In Worksheets
    If IsEmpty(Worksheet.UsedRange) Then Worksheet.Delete
    If Worksheet.Range("A2") = "" Then Worksheet.Delete
Next
Application.DisplayAlerts = True
End Function
Public Function Decir(Oracion As String)
Application.Speech.Speak Oracion, False, , True
End Function

Public Function BorrarDatosLibro()
Application.DisplayAlerts = False
For Each Worksheet In Worksheets
Worksheet.Cells.Delete
Next
Application.DisplayAlerts = True
End Function
Sub Titulos()
    'AGREGANDO TITULOS
    If ActiveSheet.Name = "Model Space" Then Call TitulosGral
    If ActiveSheet.Name = "Paper Space" Then Call TitulosGral
    If ActiveSheet.Name = "AcDbLayerTableRecord" Then Call TitulosdeLayers
    If ActiveSheet.Name = "AcDbLine-MS" Then Call TitulosdeLinea
    If ActiveSheet.Name = "AcDbLine-PS" Then Call TitulosdeLinea
    
    
    

End Sub

Public Function TitulosdeArco()


End Function

Public Function TitulosGral()
Range("A1").Value = "N"
Range("B1").Value = "TYPE"
Range("C1").Value = "NAME"
Range("A1:Z1").Interior.Color = 15773696
Range("A1").Select
End Function



Public Function TitulosdeLinea()

Range("A1").Value = "N"
Range("B1").Value = "TYPE"
Range("C1").Value = "NAME"
Range("D1").Value = "START X"
Range("E1").Value = "START Y"
Range("F1").Value = "START Z"
Range("G1").Value = "END X"
Range("H1").Value = "END Y"
Range("I1").Value = "END Z"
Range("J1").Value = "COLOR"
Range("K1").Value = "LAYER"
Range("L1").Value = ""
Range("M1").Value = ""
Range("N1").Value = ""
Range("O1").Value = ""
Range("P1").Value = ""
Range("Q1").Value = ""
Range("R1").Value = ""
Range("S1").Value = ""
Range("T1").Value = ""
Range("U1").Value = ""
Range("V1").Value = ""
Range("W1").Value = ""
Range("X1").Value = ""
Range("Y1").Value = ""
Range("Z1").Value = ""
Range("A1:Z1").Interior.Color = 15773696
Range("A1").Select

End Function

Public Function TitulosdeLayers()

Range("A1").Value = "N"
Range("B1").Value = "TYPE"
Range("C1").Value = "NAME"
Range("D1").Value = "COLOR"
Range("E1").Value = "LINETYPE"
Range("F1").Value = "LINEWEIGHT"
Range("G1").Value = "PLOTTABLE"
Range("H1").Value = ""
Range("I1").Value = ""
Range("J1").Value = ""
Range("K1").Value = ""
Range("L1").Value = ""
Range("M1").Value = ""
Range("N1").Value = ""
Range("O1").Value = ""
Range("P1").Value = ""
Range("Q1").Value = ""
Range("R1").Value = ""
Range("S1").Value = ""
Range("T1").Value = ""
Range("U1").Value = ""
Range("V1").Value = ""
Range("W1").Value = ""
Range("X1").Value = ""
Range("Y1").Value = ""
Range("Z1").Value = ""
Range("A1:Z1").Interior.Color = 15773696
Range("A1").Select

End Function
