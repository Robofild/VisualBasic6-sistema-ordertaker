Attribute VB_Name = "checkImpressa"
Option Explicit
Dim s, P As String


Public Function VerifiqueImpressão() As Boolean



s = Procura_Arquivo("C:\Windows\System32\spool\PRINTERS", "*.SHD")
P = Procura_Arquivo("C:\Windows\System32\spool\PRINTERS", "*.SPL")
If (s <> "" And P <> "") Then
MsgBox "Verifique Erro na impressora", vbCritical, "Erro ao imprimir"
VerifiqueImpressão = False
Else
MsgBox "Documento impresso com sucesso!!!", vbYes, "Sucesso ao imprimir"
VerifiqueImpressão = True

End If

End Function
Public Function Procura_Arquivo(Caminho As String, NomeArquivo As String) As String

Dim lNullPos As Long
Dim lResultado As Long
Dim sBuffer As String

On Error GoTo Procura_Arquivo_Error

'Aloca espaco para a string sBuffer
sBuffer = Space(MAX_PATH * 2)
'inicia busca do arquivo
lResultado = SearchTreeForFile(Caminho, NomeArquivo, sBuffer)

' Se houver um caracter Nulo , remove
If lResultado Then
   lNullPos = InStr(sBuffer, vbNullChar)
    If Not lNullPos Then
       sBuffer = Left(sBuffer, lNullPos - 1)
    End If
   'Retorna o nome do arquivo encontrado
    Procura_Arquivo = sBuffer
Else
    'nao achou nada
    Procura_Arquivo = vbNullString
End If

Exit Function
Procura_Arquivo_Error:
    Procura_Arquivo = vbNullString
End Function
