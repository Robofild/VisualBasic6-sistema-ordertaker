Attribute VB_Name = "queryunloadme"
Option Explicit

Dim alteracaoReg As Integer
Dim respostaUser As Integer



Function FecharSairSalvaar(nomeform As String) As Integer


If (alteracaoReg <> 0) Then
respostaUser = MsgBox("Voc� n�o salvou as altera��es no : " & nomeform & " deseja fechar mesmo assim", vbCritical + vbOKCancel, "N�o salvar as altera��es")
 
 
FecharSairSalvaar = respostaUser
Else

 FecharSairSalvaar = 1


End If

End Function

Function Alterar(fechar) As Integer
alteracaoReg = fechar

End Function
