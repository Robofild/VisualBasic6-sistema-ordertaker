Attribute VB_Name = "ImpressaoCancelar"
Option Explicit


Public Sub CANCELAMENTOs(numPedido As Integer)
'CANCELAR PEDIDO
ConServer

Dim NpTitulo As String

Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs.ActiveConnection = con
rs.CursorLocation = adUseClient


sql = "UPDATE `at_contadorDePedidos` SET `situacaoImpressao` = '0' WHERE `at_contadorDePedidos`.`id` =  '" & numPedido & "'"
                            rs.Open sql
Printer.Print " -----------------------------------------------"
Printer.Print "CANCELAMENTO DO PEDIDO:  ", numPedido
Printer.Print ")"
Printer.Print ""
Printer.Print ""
Printer.Print " -----------------------------------------------"
Printer.Print "PEDIDO CANCELADO DURANTE A OPERAÇÃO"
Printer.Print " -----------------------------------------------"
Printer.Print "DESCONSIDERAR A FALTA DESTE PEDIDO  "
Printer.Print " -----------------------------------------------"
Printer.Print " "
 Set rs = Nothing
End Sub

