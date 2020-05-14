Attribute VB_Name = "imagem"
Option Explicit
Private Const nBUFFER As Long = 1024
'imagens normais
Public Sub SalvaImagem(f As ADODB.Field, File As String)
    Dim b() As Byte
    Dim ff  As Long
    Dim n   As Long
    
    On Error GoTo ErrHandler
    ff = FreeFile
    Open File For Binary Access Read As ff
    n = LOF(ff)
    If n Then
       ReDim b(1 To n) As Byte
       Get ff, , b()
    End If
    Close ff
    f.Value = b()
    Exit Sub
    
ErrHandler:
    MsgBox "ERROR: " & Err.Description
End Sub
Public Function RecuperaImagem(f As ADODB.Field) As StdPicture
    
    Dim b()  As Byte
    Dim ff   As Long
    Dim File As String
    
    On Error GoTo ErrHandler
    Call GetRandomFileName(File)
    ff = FreeFile
    Open File For Binary Access Write As ff
    b() = f.Value
    Put ff, , b()
    Close ff
    Erase b
    Set GetImageFromField = LoadPicture(File)
    Kill File
    Exit Function
    
ErrHandler:
    MsgBox "ERROR: " & Err.Description
End Function

'imagens grandes
Public Sub SalvaImagensGrandes(f As ADODB.Field, File As String)
    Dim b()      As Byte
    Dim ff       As Long
    Dim i        As Long
    Dim FileLen  As Long
    Dim Blocks   As Long
    Dim LeftOver As Long

    On Error GoTo ErrHandler
    ff = FreeFile
    Open File For Binary Access Read As ff

    FileLen = LOF(ff)
    Blocks = Int(FileLen / nBUFFER)
    LeftOver = FileLen Mod nBUFFER
  
    ReDim b(LeftOver)
    Get ff, , b()
    f.AppendChunk b()
  
    ReDim b(nBUFFER)
    For i = 1 To Blocks
        Get ff, , b()
        f.AppendChunk b()
    Next
    Close ff
    Exit Sub
    
ErrHandler:
    MsgBox "ERROR: " & Err.Description
End Sub
Public Function ExibeImagensGrandes(f As ADODB.Field) As StdPicture
    
    Dim b()      As Byte
    Dim ff       As Long
    Dim File     As String
    Dim i        As Long
    Dim FileLen  As Long
    Dim Blocks   As Long
    Dim LeftOver As Long

    On Error GoTo ErrHandler
    File = "temppic.bmp"
    ff = FreeFile
    Open File For Binary Access Write As ff
    Blocks = Int(f.ActualSize / nBUFFER)
    LeftOver = f.ActualSize Mod nBUFFER
    b() = f.GetChunk(LeftOver)
    Put ff, , b()
    For i = 1 To Blocks
        b() = f.GetChunk(nBUFFER)
        Put ff, , b()
    Next
    Close ff
    Erase b
    Set ExibeImagensGrandes = LoadPicture(File)
    Kill File
    Exit Function
    
ErrHandler:
    MsgBox "ERROR: " & Err.Description
End Function
