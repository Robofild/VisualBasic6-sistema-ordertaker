Attribute VB_Name = "Devolucaodocardapio"
Option Explicit
Dim indice As Integer


Public Function Devolucao(DevolverPara As String) As Integer
'trarar erro
On Error GoTo error

indice = Form52.DataGrid1.Columns(0).Value



Select Case DevolverPara
Case 50
 'retornar para fomulario cadastro de cardapio compor igredientes
 
 Form50.Text7.Text = indice
 Form50.CmdExcluir.Enabled = True
 

 
    


 
 
 
End Select

Exit Function

error:
MsgBox "Não foi possivel definir o produto, por favor continue a escolher ", vbYes, "Defina melhor o produto"
Exit Function


End Function
