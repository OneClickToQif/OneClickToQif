Attribute VB_Name = "modulo_exportar"

Sub ExportAllToQIF()

    Dim formulario As Range
    Set formulario = Range("b4")
    FolderName = formulario.Cells(1, 2).Value
    
   Dim WS_Count As Integer
   Dim I As Integer

   ' Set WS_Count equal to the number of worksheets in the active workbook.
   WS_Count = ActiveWorkbook.Worksheets.Count

    ' cojo la variable global ya establecida
    ' Dim FolderName As String
    ' FolderName = GetFolderName("Select a folder")
    If FolderName = "" Then
        MsgBox "You didn't select a folder."
    Else
           For I = 2 To WS_Count   ' la primera es la hoja de control
     
              formulario.Cells(I + 3, 1).Value = "-"
              formulario.Cells(I + 3, 2).Value = "-"
        
           Next I
           For I = 2 To WS_Count   ' la primera es la hoja de control
                Dim mensaje As String
                Dim nombre As String

              nombre = ActiveWorkbook.Worksheets(I).Name
              mensaje = ExportToQIF(FolderName & "\" & nombre, I)
              
              formulario.Cells(I + 3, 1).Value = nombre
              formulario.Cells(I + 3, 2).Value = mensaje
        
           Next I
           ThisWorkbook.Worksheets(1).Select
    End If

End Sub

Public Function ExportToQIF(FullPath As String, NumHoja As Integer) As String


On Error GoTo ErrorHandler

Dim oWorkSheet As Worksheet
Dim lRows As Long
Dim lCols As Long
Dim iFileNum As Integer
Dim iFileNum2 As Integer

Dim numExportadas As Long

Dim ColumnaCategory As Integer
Dim ColumnaAmount As Integer
Dim ColumnaCategory2 As Integer
Dim ColumnaAmount2 As Integer
Dim ColumnaDate As Integer
Dim ColumnaMemo As Integer
Dim ColumnaExportado As Integer

ColumnaCategory = NumeroDeColumna(NumHoja, "(Category)")
ColumnaAmount = NumeroDeColumna(NumHoja, "(Amount)")
ColumnaCategory2 = NumeroDeColumna(NumHoja, "(Category2)")
ColumnaAmount2 = NumeroDeColumna(NumHoja, "(Amount2)")
ColumnaDate = NumeroDeColumna(NumHoja, "(Date)")
ColumnaMemo = NumeroDeColumna(NumHoja, "(Memo)")
ColumnaExportado = NumeroDeColumna(NumHoja, "(Exported)")

' MsgBox ("aqui " & "-" & ColumnaCategory & "-" & ColumnaMemo & "-" & ColumnaDate & "-" & ColumnaAmount & "-" & ColumnaExportado)

If ColumnaExportado = -1 Then
    ExportToQIF = "Error: Columna Exported not found (Maybe this is not a data tab) - No encontre columna Exported"
    Exit Function
End If
If ColumnaCategory = -1 Then
    ExportToQIF = "Error: ColumnaCategory not found"
    Exit Function
End If
If ColumnaDate = -1 Then
    ExportToQIF = "Error: ColumnaDate not found"
    Exit Function
End If
If ColumnaMemo = -1 Then
    ExportToQIF = "Error: ColumnaMemo not found"
    Exit Function
End If
If ColumnaAmount = -1 Then
    ExportToQIF = "Error: ColumnaAmount not found"
    Exit Function
End If

Set oWorkSheet = ThisWorkbook.Worksheets(NumHoja)
sName = oWorkSheet.Name
lCols = oWorkSheet.UsedRange.Columns.Count
lRows = oWorkSheet.UsedRange.Rows.Count
oWorkSheet.Select

' primero miramos que haya algo que exportar
numExportadas = 0
For I = 3 To lRows
    'If Trim(Cells(I, ColumnaExportar).Value) = "" Then
    '    Exit For
    'End If
    
    If UCase(Trim(Cells(I, ColumnaExportado).Value)) = "N" Then
        numExportadas = numExportadas + 1
    End If
    
Next I

If numExportadas = 0 Then
    ExportToQIF = "No rows to export [No tiene filas por exportar]"
    Exit Function
End If

For I = 3 To lRows
    If UCase(Trim(Cells(I, ColumnaExportado).Value)) = "N" Then
        If Cells(I, ColumnaDate).Value = "" Then
            ExportToQIF = "Date column is empty for row [Columna fecha vacia para fila]: " & I
            Exit Function
        End If
        If Cells(I, ColumnaAmount).Value = "" Then
            If ColumnaAmount2 = -1 Then ' si hay amount2 perdono que no haya amount
                ExportToQIF = "Amount column is empty for row [Columna amount vacia para fila]: " & I
                Exit Function
            End If
        End If
        If Cells(I, ColumnaMemo).Value = "" Then
            ExportToQIF = "Memo column es empty for row [Columna memo vacia para fila]: " & I
            Exit Function
        End If
        If Cells(I, ColumnaCategory).Value = "" Then
            If ColumnaAmount2 = -1 Then ' si hay amount2 perdono que no haya amount ni category
                ExportToQIF = "Category column is empty for row [Columna category vacia para fila]: " & I
                Exit Function
            End If
        End If
        
        If ColumnaAmount2 <> -1 Then
            If Cells(I, ColumnaAmount2).Value <> "" And Cells(I, ColumnaCategory2).Value = "" Then
                ExportToQIF = "Category2 column is empty for row [Columna category2 vacia para fila]: " & I
                Exit Function
            End If
        End If
    End If
    
Next I

If Dir(FullPath & ".qif") > "" Then
    ExportToQIF = "File already exists [El fichero ya existe]: " & FullPath
    Exit Function
End If

If Dir(FullPath & "_2.qif") > "" Then
    ExportToQIF = "File already exists [El fichero ya existe]: " & FullPath
    Exit Function
End If

' abrimos el fichero para escribir en el
iFileNum = FreeFile
Open FullPath & ".qif" For Output As #iFileNum
Print #iFileNum, "!Type:Bank"

If ColumnaAmount2 <> -1 Then
    iFileNum2 = FreeFile
    Open FullPath & "_2.qif" For Output As #iFileNum2
    Print #iFileNum2, "!Type:Bank"
End If

Application.ScreenUpdating = False

numExportadas = 0
For I = 3 To lRows
    If UCase(Trim(Cells(I, ColumnaExportado).Value)) = "N" Then
        numExportadas = numExportadas + 1
        Cells(I, ColumnaExportado).Value = "Y"
        If Cells(I, ColumnaAmount).Value <> "" Then ' podria darse el caso de que estuviera vacio, si hay amount2 perdono que no haya amount
            Print #iFileNum, "D" & Cells(I, ColumnaDate).Value
            Print #iFileNum, "U" & Replace(Cells(I, ColumnaAmount).Value, ",", ".")
            Print #iFileNum, "T" & Replace(Cells(I, ColumnaAmount).Value, ",", ".")
            Print #iFileNum, "M" & Cells(I, ColumnaMemo).Value
            Print #iFileNum, "L" & Cells(I, ColumnaCategory).Value
            Print #iFileNum, "^"
        End If
        If ColumnaAmount2 <> -1 Then
            If Cells(I, ColumnaAmount2).Value <> "" Then
                Print #iFileNum2, "D" & Cells(I, ColumnaDate).Value
                Print #iFileNum2, "U" & Replace(Cells(I, ColumnaAmount2).Value, ",", ".")
                Print #iFileNum2, "T" & Replace(Cells(I, ColumnaAmount2).Value, ",", ".")
                Print #iFileNum2, "M" & Cells(I, ColumnaMemo).Value
                Print #iFileNum2, "L" & Cells(I, ColumnaCategory2).Value
                Print #iFileNum2, "^"
            End If
        End If
    End If

Next I

Application.ScreenUpdating = True

Close #iFileNum
If ColumnaAmount2 = -1 Then
    ExportToQIF = numExportadas & " rows have been exported [Exportadas " & numExportadas & " filas]."
Else
    Close #iFileNum2
    ExportToQIF = "Exportadas " & numExportadas & " filas a 2 ficheros."
End If
Exit Function

ErrorHandler:
    Close #iFileNum
    ExportToQIF = "Error exporting sheet [Error en hoja]"

Exit Function
End Function


Public Function NumeroDeColumna(NumeroHoja As Integer, NombreColumna As String) As Integer


On Error GoTo ErrorHandler

Dim lCols As Long
Dim oWorkSheet As Worksheet

Set oWorkSheet = ThisWorkbook.Worksheets(NumeroHoja)
lCols = oWorkSheet.UsedRange.Columns.Count
oWorkSheet.Select

Dim encontrado As Boolean
encontrado = False


For I = 1 To lCols
    'presupongo fila 2
    If Trim(Cells(2, I).Value) = NombreColumna Then encontrado = True
    If encontrado Then Exit For
Next I

If Not encontrado Then GoTo ErrorHandler

NumeroDeColumna = I
Exit Function

ErrorHandler:
    
    NumeroDeColumna = -1

Exit Function
End Function



