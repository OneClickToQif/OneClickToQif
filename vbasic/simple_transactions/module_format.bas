Attribute VB_Name = "modulo_formatos"
Sub dar_formato()
Attribute dar_formato.VB_ProcData.VB_Invoke_Func = " \n14"

    Sheets("formatos").Visible = True

    Call dar_formato_tab(Sheets("cash"))
    
    Columns("E:F").EntireColumn.Hidden = True
    Columns("H").ColumnWidth = 0.75
    Columns("J:L").EntireColumn.Hidden = True
    
    Call dar_formato_tab(Sheets("checking_account"))
    Call dar_formato_tab(Sheets("saving_account"))
    Call dar_formato_tab(Sheets("credit_card"))
    
    Sheets("formatos").Visible = False
   
End Sub
Function dar_formato_tab(hoja As Worksheet)

    hoja.Select
    Cells.FormatConditions.Delete
    Cells.ClearFormats
    
    Sheets("formatos").Select
    Cells.Select
    Selection.Copy
    
    hoja.Select
    Cells.Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    Cells(1, 1).Select

End Function

