Attribute VB_Name = "Módulo1"
Option Explicit
Function PolyCoords(oEnt As AcadEntity) As Variant
    Dim cnt As Integer
    Dim i As Integer
    Dim j As Integer
    Dim iStep As Integer
    Dim varPt As Variant
    Dim dblCoords() As Double
    Dim dblVert() As Double

    If TypeOf oEnt Is AcadLWPolyline Then
         iStep = 2
    ElseIf TypeOf oEnt Is Acad3DPolyline Or _
           TypeOf oEnt Is AcadPolyline Then
         iStep = 3
    End If
    dblCoords = oEnt.Coordinates

    ReDim ptsArr(0 To (UBound(dblCoords) + 1) \ iStep - 1, 0 To iStep - 1) As Double
    For i = 0 To (UBound(dblCoords) + 1) \ iStep - 1
         For j = 0 To iStep - 1
              ptsArr(i, j) = dblCoords(cnt)
              Debug.Print ptsArr(i, j)
              cnt = cnt + 1
         Next
    Next
    PolyCoords = ptsArr
End Function

Sub demo()
    Dim pts As Variant
    Dim varPt As Variant
    Dim oEnt As AcadEntity
    AutoCAD.Application.ActiveDocument.Utility.GetEntity oEnt, varPt, vbCr & "Select polyline"
    If Not TypeOf oEnt Is AcadLWPolyline And _
       Not TypeOf oEnt Is Acad3DPolyline And _
       Not TypeOf oEnt Is AcadPolyline Then
         MsgBox "Method is not applicable for this entity type"
         Exit Sub
    End If
    pts = PolyCoords(oEnt)
End Sub



Sub Example_Coordinates()

Dim Selection As AcadSelectionSet
Dim Poly As AcadLWPolyline
Dim Obj As AcadEntity
Dim Bound As Double
Dim x As Long
Dim Y As Long
Dim i As Long
Dim j As Long

'Makes a selectionset.
On Error Resume Next
   Set Selection = AutoCAD.Application.ActiveDocument.SelectionSets.Item("Select polyline.")
If Err Then
   Set Selection = AutoCAD.Application.ActiveDocument.SelectionSets.Add("Select polyline.")
   Err.Clear
Else
   Selection.Clear
End If

'Select the polyline.
Selection.SelectOnScreen

For Each Obj In Selection

   If Obj.ObjectName = "AcDbPolyline" Then
       
           Set Poly = Obj
           On Error Resume Next
           
           Bound = UBound(Poly.Coordinates)
           
           x = 0
           Y = 1
           
           For i = 0 To Bound / 2
               
               
               
               MsgBox "X= " & Poly.Coordinates(x) & "  Y= " & Poly.Coordinates(Y)
               If Err Then Err.Clear
               
               x = x + 2
               Y = Y + 2
               
           Next
         
   End If

Next Obj

End Sub
