Attribute VB_Name = "Ordenar_AZ"
Sub OrdenarAZ()
Attribute OrdenarAZ.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Ordenar_AZ Macro
'

'
    ActiveWorkbook.Worksheets("Lista Coletas").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Lista Coletas").Sort.SortFields.Add2 Key:=Range( _
        "B3:B27"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Lista Coletas").Sort
        .SetRange Range("B3:I27")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
