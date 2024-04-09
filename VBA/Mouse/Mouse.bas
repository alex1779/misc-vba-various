Attribute VB_Name = "Mouse"
Public Declare PtrSafe Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As LongPtr

Public Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Type POINTAPI
   Xcoord As Long
   Ycoord As Long
End Type
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare PtrSafe Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Public Const MOUSEEVENTF_LEFTDOWN = &H2
Public Const MOUSEEVENTF_LEFTUP = &H4
Public Const MOUSEEVENTF_RIGHTDOWN As Long = &H8
Public Const MOUSEEVENTF_RIGHTUP As Long = &H10
'......................

 Sub LeftClick()
  mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
  Sleep 50
  mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
End Sub
 Sub RightClick()
  mouse_event MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0
  Sleep 50
  mouse_event MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0
End Sub

Sub MinimizarVBA()
SetCursorPos 1250, 10
LeftClick
Sleep 2000
End Sub

Sub Nuevo_Procedimiento_Click()
Dim llCoord As POINTAPI
GetCursorPos llCoord
titulo = Application.InputBox("Titulo del Procedimiento")
MinimizarVBA
Dim FilePath As String
Dim TextFile As Integer
RutaArchivo = Application.ActiveWorkbook.Path & "\" & titulo & ".txt"
Open RutaArchivo For Output As #1
Print #1, "Sub " & titulo & "()"
Print #1, "SetCursorPos " & llCoord.Xcoord & ", " & llCoord.Ycoord
Print #1, "LeftClick"
Print #1, "'Sleep 2000"
Print #1, "End Sub"
Close #1
abrirtxt = Shell("notepad.exe " & RutaArchivo, vbMaximizedFocus)
'Copiar_TXT
End Sub