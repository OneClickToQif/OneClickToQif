Attribute VB_Name = "Module1"
Private Type BROWSEINFO ' used by the function GetFolderName
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
    
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long

Dim FolderName As String
    
Sub seleccionaCarpeta()
    Dim formulario As Range
    Set formulario = Range("h1")
    
    FolderName = GetFolderName("Select a folder")
    If FolderName = "" Then
        MsgBox "You didn't select a folder."
    Else
        formulario.Cells(1, 1).Value = FolderName
    End If

End Sub

Public Function ValoresDeFila(oWorkSheet As Worksheet, NombrePrimeraColumna As String) As String()

    On Error GoTo ErrorHandler

    Dim lCols As Long
    Dim lRows As Long

    lCols = oWorkSheet.UsedRange.Columns.Count
    lRows = oWorkSheet.UsedRange.Rows.Count
    oWorkSheet.Select

    Dim encontrado As Boolean
    encontrado = False
    
    For j = 1 To lRows
        If Trim(Cells(j, 1).Value) = NombrePrimeraColumna Then encontrado = True
        If encontrado Then Exit For
    Next j

    If Not encontrado Then GoTo ErrorHandler

    Dim numTitulos As Integer
    numTitulos = 0

    For i = 2 To lCols
        If Trim(Cells(j, i).Value) <> "" Then numTitulos = numTitulos + 1
    Next i

    Dim valores() As String
    ReDim valores(numTitulos - 1)
    numTitulos = 0
    For i = 2 To lCols
        If Trim(Cells(j, i).Value) <> "" Then
            valores(numTitulos) = Trim(Cells(j, i).Value)
            numTitulos = numTitulos + 1
        End If
    Next i

    ValoresDeFila = valores
Exit Function

ErrorHandler:
    MsgBox ("Not found in this tab [No se ha encontrado] " & NombrePrimeraColumna & " [en la hoja]")
    
Exit Function

End Function


Public Function esColumnaEstandar(nombreCol As String) As Boolean

    Dim nombre As String
    nombre = Normaliza(nombreCol)
    
    If nombre = "EXPORTED" Or nombre = "DATE" Or nombre = "MAIN" Or nombre = "MEMO" Then
        esColumnaEstandar = True
    Else
        esColumnaEstandar = False
    End If

End Function

Public Function Normaliza(nombre As String) As String

    Dim normalizada As String
    
    With Application.WorksheetFunction
    normalizada = .Substitute(nombre, "-", "")
    normalizada = .Substitute(normalizada, " ", "")
    ' normalizada = .Substitute(normalizada, char(160), "") ' non breaking space
    normalizada = StrConv(normalizada, vbUpperCase)

    End With
    Normaliza = normalizada
    
    
End Function


Public Function NumeroDeColumna(oWorkSheet As Worksheet, NombreColumna As String) As Integer


On Error GoTo ErrorHandler

Dim lCols As Long
Dim lRows As Long

' Dim oWorkSheet As Worksheet
' Set oWorkSheet = ThisWorkbook.Worksheets(NumeroHoja)

lCols = oWorkSheet.UsedRange.Columns.Count
lRows = oWorkSheet.UsedRange.Rows.Count
oWorkSheet.Select

Dim encontrado As Boolean
encontrado = False

For j = 1 To lRows
    If Trim(Cells(j, 1).Value) = "QIF" Then encontrado = True
    If encontrado Then Exit For
Next j

If Not encontrado Then GoTo ErrorHandler

encontrado = False

For i = 2 To lCols
    If Normaliza(Cells(j, i).Value) = Normaliza(NombreColumna) Then encontrado = True
    If encontrado Then Exit For
Next i

If Not encontrado Then GoTo ErrorHandler

NumeroDeColumna = i
Exit Function

ErrorHandler:
    ' MsgBox ("No se ha encontrado " & NombreColumna & " en hoja numero " & NumeroHoja)
    NumeroDeColumna = -1

Exit Function
End Function


Function GetFolderName(Msg As String) As String
' returns the name of the folder selected by the user
Dim bInfo As BROWSEINFO, path As String, r As Long
Dim X As Long, pos As Integer
    bInfo.pidlRoot = 0& ' Root folder = Desktop
    If IsMissing(Msg) Then
        bInfo.lpszTitle = "Select a folder."
        ' the dialog title
    Else
        bInfo.lpszTitle = Msg ' the dialog title
    End If
    bInfo.ulFlags = &H1 ' Type of directory to return
    X = SHBrowseForFolder(bInfo) ' display the dialog
    ' Parse the result
    path = Space$(512)
    r = SHGetPathFromIDList(ByVal X, ByVal path)
    If r Then
        pos = InStr(path, Chr$(0))
        GetFolderName = Left(path, pos - 1)
    Else
        GetFolderName = ""
    End If
End Function

