Attribute VB_Name = "Módulo4"
Sub Macro4()
Attribute Macro4.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro4 Macro
'

'
    ActiveWorkbook.Worksheets("AcDbLine-MS").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("AcDbLine-MS").Sort.SortFields.Add Key:=Range("D2") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("AcDbLine-MS").Sort
        .SetRange Range("D2:E32")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
