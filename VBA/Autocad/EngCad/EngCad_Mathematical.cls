VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EngCad_Mathematical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function Random(menor As Integer, mayor As Integer)
Random = ((mayor - menor + 1) * Rnd + menor)
Do While Random = 0: Random = ((mayor - menor + 1) * Rnd + menor): Loop 'Non Zero
Do While Random > mayor: Random = ((mayor - menor + 1) * Rnd + menor): Loop 'Non Zero
End Function

