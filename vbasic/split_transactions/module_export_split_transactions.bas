Attribute VB_Name = "Module2"
Sub ExportHoja()
  
    FolderName = Range("carpeta_exportacion_qif").Value
    If FolderName = "" Then
        MsgBox "You didn't select a folder [No has seleccionado la carpeta]"
        Exit Sub
    End If
    
    Dim nombre As String
    nombre = ActiveSheet.Name
    If nombre = "" Then
        MsgBox "Write a file name [Escribe un nombre de fichero]"
        Exit Sub
    End If

    mensaje = ExportHojaToQIF(FolderName & "\" & nombre & ".qif", ActiveSheet)
    
    ' formulario.Cells(1, 10).Value = mensaje
    MsgBox (mensaje)

End Sub


Public Function ExportHojaToQIF(FullPath As String, oWorkSheet As Worksheet) As String


On Error GoTo ErrorHandler

' Dim oWorkSheet As Worksheet
Dim lRows As Long
Dim lCols As Long
Dim iFileNum As Integer

' Set oWorkSheet = ThisWorkbook.Worksheets(NumHoja)
' Set oWorkSheet = ActiveSheet
sName = oWorkSheet.Name
lCols = oWorkSheet.UsedRange.Columns.Count
lRows = oWorkSheet.UsedRange.Rows.Count
oWorkSheet.Select

Dim numExportadas As Long

Dim ColDate As Integer
Dim ColMemo As Integer
Dim ColExportado As Integer
Dim colPpalAsiento As Integer

Dim columnasExportables() As String
columnasExportables = ValoresDeFila(oWorkSheet, "QIF")

ColDate = NumeroDeColumna(oWorkSheet, "date")
If ColDate = -1 Then
    ExportHojaToQIF = "Error: Date column not found [no encontre columna Date]"
    Exit Function
End If

ColMemo = NumeroDeColumna(oWorkSheet, "memo")
If ColMemo = -1 Then
    ExportHojaToQIF = "Error: Memo column not found [no encontre columna Memo]"
    Exit Function
End If

ColExportado = NumeroDeColumna(oWorkSheet, "exported")
If ColExportado = -1 Then
    ExportHojaToQIF = "Error: Exported column not found [no encontre columna Exported]"
    Exit Function
End If

colPpalAsiento = NumeroDeColumna(oWorkSheet, "main")
If colPpalAsiento = -1 Then
    ExportHojaToQIF = "Error: Main column not found [no encontre columna Main]"
    Exit Function
End If

' primero miramos que haya algo que exportar
numExportadas = 0
For i = 1 To lRows
    If UCase(Trim(Cells(i, ColExportado).Value)) = "N" Then
        numExportadas = numExportadas + 1
    End If
    
Next i

If numExportadas = 0 Then
    ExportHojaToQIF = "No rows to export [No tiene filas por exportar]"
    Exit Function
End If

For i = 1 To lRows
    
    If UCase(Trim(Cells(i, ColExportado).Value)) = "N" Then
        If Cells(i, ColDate).Value = "" Then
            ExportHojaToQIF = "Date column is empty for row [Columna fecha esta vacia para fila]: " & i
            Exit Function
        End If
        If Cells(i, ColMemo).Value = "" Then
            ExportHojaToQIF = "column is empty for row [Columna memo vacia para fila]:" & i
            Exit Function
        End If
        If Cells(i, colPpalAsiento).Value = "" Then
            ExportHojaToQIF = "Main column is empty for row [Columna Main vacia para fila]: " & i
            Exit Function
        End If
    End If
Next i

If Dir(FullPath) > "" Then
    ExportHojaToQIF = "File already exists: " & FullPath & "[El fichero ya existe]"
    Exit Function
End If

' abrimos el fichero para escribir en el
iFileNum = FreeFile
Open FullPath For Output As #iFileNum

Print #iFileNum, "!Type:Bank"

Application.ScreenUpdating = False

numExportadas = 0
For i = 1 To lRows
  
    If UCase(Trim(Cells(i, ColExportado).Value)) = "N" Then
        numExportadas = numExportadas + 1
        Print #iFileNum, "D" & Cells(i, ColDate).Value
        Dim valor As Double
        valor = Cells(i, colPpalAsiento).Value
        If valor = 0 Then  ' hay un error en gnucash cuando es 0, importa con valor contrario la siguiente fila
                            ' de todos modos esto no se usa al importar para asientos de multiples entradas
            valor = 0.01
        End If
        
        Print #iFileNum, "U" & Replace(valor, ",", ".")
        Print #iFileNum, "T" & Replace(valor, ",", ".")
        Print #iFileNum, "M" & Cells(i, ColMemo).Value & " - asiento"
        Print #iFileNum, "L" & "asiento"
        
        For k = 0 To UBound(columnasExportables)
            Dim titulo As String
            titulo = columnasExportables(k)
            
            If Not esColumnaEstandar(titulo) Then
                Dim columna As Integer
                columna = NumeroDeColumna(oWorkSheet, titulo)
                If Cells(i, columna).Value <> "" And Cells(i, columna).Value <> 0 Then
                    valor = Cells(i, columna).Value
                    If InStr(1, titulo, "(neg)") Then
                        valor = -valor
                        titulo = Replace(titulo, "(neg)", "")
                        titulo = Trim(titulo)
                    End If
                    Print #iFileNum, "S" & titulo
                    Print #iFileNum, "E" & Cells(i, ColMemo).Value & " - " & titulo
                    Print #iFileNum, "$" & Replace(valor, ",", ".")
                End If
            End If
        Next k

        Print #iFileNum, "^"
        Cells(i, ColExportado).Value = "Y"
    End If

Next i

Application.ScreenUpdating = True

Close #iFileNum
ExportHojaToQIF = "" & numExportadas & " rows exported [filas exportadas]."
Exit Function

ErrorHandler:
    Close #iFileNum
    ExportHojaToQIF = "Error at tab [Error en hoja] : " + Error

Exit Function
End Function

