Attribute VB_Name = "queryunloadme"
Option Explicit

Dim alteracaoReg As Integer
Dim respostaUser As Integer



Function FecharSairSalvaar(nomeform As String) As Integer


If (alteracaoReg <> 0) Then
respostaUser = MsgBox("Você não salvou as alterações no : " & nomeform & " deseja fechar mesmo assim", vbCritical + vbOKCancel, "Não salvar as alterações")
 
 
FecharSairSalvaar = respostaUser
Else

 FecharSairSalvaar = 1


End If

End Function

Function Alterar(fechar) As Integer
alteracaoReg = fechar

End Function
