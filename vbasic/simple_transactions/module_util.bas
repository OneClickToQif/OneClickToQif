Attribute VB_Name = "Modulo_utils"
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
    Set formulario = Range("b4")
    
    
    FolderName = GetFolderName("Select a folder")
    If FolderName = "" Then
        MsgBox "You didn't select a folder."
    Else
        formulario.Cells(1, 2).Value = FolderName
    End If

End Sub

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

Sub FlipRows()
    Dim vTop As Variant
    Dim vEnd As Variant
    Dim iStart As Integer
    Dim iEnd As Integer
        Application.ScreenUpdating = False
        iStart = 1
        iEnd = Selection.Rows.Count
        Do While iStart < iEnd
            vTop = Selection.Rows(iStart)
            vEnd = Selection.Rows(iEnd)
            Selection.Rows(iEnd) = vTop
            Selection.Rows(iStart) = vEnd
            iStart = iStart + 1
            iEnd = iEnd - 1
        Loop
        Application.ScreenUpdating = True
End Sub


