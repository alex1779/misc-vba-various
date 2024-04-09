Attribute VB_Name = "Módulo1"
Public Function RecortayPega()

    Dim n As Integer
    n = Application.WorksheetFunction.CountA(Range("A1:A2000"))
    Dim ultimafila As Integer
    Dim FirstPos As Integer
    Dim EachCell As Integer
    FirstPos = Application.InputBox(Prompt:="PRIMERA POSICION", Type:=2)
    EachCell = Application.InputBox(Prompt:="REPETICION", Type:=2)
    a = 1
    Do While FirstPos <= n
        Cells(FirstPos, 1).Select
        Selection.Cut
        Cells(a, 1).Select
        ActiveSheet.Paste
    FirstPos = FirstPos + EachCell
    a = a + 1
    Loop
 
    MsgBox FirstPos
End Function
Public Function BorraFilasEnBlanco()
'Macro Para Borrar las Celdas en Blanco dentro de la columna A
Dim i As Integer
    For i = 1 To 200
    If Cells(i, 1).Value = "" Then
    Cells(i, 1).Select
    Selection.Delete
    End If
Next i
Cells(1, 1).Select
End Function
Public Function Reemplazar(Texto1 As String, Texto2 As String)
Cells.Replace What:=Texto1, Replacement:=Texto2, LookAt:=xlPart, _
SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
ReplaceFormat:=False
End Function
Sub RemoveDuplicates()
Range("A1:A100").RemoveDuplicates Columns:=Array(1), Header:=xlYes
End Sub
Public Sub GetTranslation()
    Dim Navegador As New InternetExplorer
    Dim ws As Worksheet
    Dim inputString As String, outputString As String, text_to_convert As String, translation As String
    Dim NET As SHDocVw.InternetExplorer
    Dim Pagina As MSHTML.HTMLDocument
    Dim EntradaNet As MSHTML.HTMLInputElement
    Dim ELEMENTLAR As MSHTML.IHTMLElementCollection
    
    inputString = "en"
    outputString = "es"
    text_to_convert = Cells(1, 1)
    
    Set ws = ThisWorkbook.Worksheets("Hoja1")
        Set NET = New SHDocVw.InternetExplorer
        NET.navigate "https://translate.google.com/#" & inputString & "/" & outputString & "/" & text_to_convert
        NET.Visible = 0
        Do While NET.readyState <> 4: DoEvents: Loop
        Application.Wait (Now + TimeValue("0:00:01"))
        Set Pagina = NET.document
        Set ELEMENTLAR = Pagina.getElementsByClassName("tlid-translation translation")
        Do While Pagina.readyState <> "complete": DoEvents: Loop
            For Each EntradaNet In ELEMENTLAR
            If EntradaNet.ID = "" Then
            Cells(1, 2) = EntradaNet.innerText
            Exit For
        End If
        Next EntradaNet
        ''NET.Quit
End Sub

