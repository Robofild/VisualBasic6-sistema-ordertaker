Attribute VB_Name = "RetirarCaracteresEspeciais"
Option Explicit

Public Function RemoverCaracter(Valor As String) As String
Dim Remover As String, i As Byte, Temp As String
Remover = ".()*/-+^´"
Temp = Valor
For i = 1 To Len(Valor)
    Temp = Replace(Temp, Mid(Remover, i, 1), "")
Next
RemoverCaracter = Temp
End Function
Public Function retiraCaracteresEspeciais(strAFiltrar As String)

    Dim posASubstituir  As Integer
    Dim curPos          As Integer
    Dim curChar         As String
    Dim substituirDe    As String
    Dim substituirPara  As String
    Dim strFiltrada     As String
    
      substituirDe = "äáàãâÄÁÀéèêëËÉÈÊíìïÍÌÎÏóòôõöÓÒÔÕÖúùûüÚÙÛÜçÇ"
    substituirPara = "aaaaaAAAeeeeEEEEiiiIIIIoooooOOOOOuuuuUUUUcC"
    
    For curPos = 1 To Len(strAFiltrar) 'ciclo na string a filtrar...
        curChar = Mid(strAFiltrar, curPos, 1) 'pega em cada caracter da string
        posASubstituir = InStr(substituirDe, curChar) 'verifica se está na string de caracteres a substituir
        If posASubstituir Then  ' se estiver,
            strFiltrada = strFiltrada & Mid(substituirPara, posASubstituir, 1) 'entra na string filtrada o equivalente não-acentuado
        Else 'se não estiver
            strFiltrada = strFiltrada & curChar 'entra na string filtrada tal como está
        End If
    Next curPos
    
    retiraCaracteresEspeciais = strFiltrada
End Function
