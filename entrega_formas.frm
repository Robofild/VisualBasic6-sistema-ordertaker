VERSION 5.00
Begin VB.Form Form406 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Forma de Entrega"
   ClientHeight    =   2715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4185
   Icon            =   "entrega_formas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   4185
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   6120
      TabIndex        =   6
      Text            =   "0"
      Top             =   840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "F1-     ENTREGAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "F2-           BALCÃO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   3735
   End
   Begin VB.Label Entregabusca 
      Caption         =   "entrega busca"
      Height          =   375
      Left            =   6120
      TabIndex        =   7
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   4440
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "4.75"
      Height          =   375
      Left            =   4560
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "Form406"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  Form407.Label10.Caption = 1
  Form404.Text11.Text = 1

 removerfrete
'BALCAO

Form407.Show
Form407.Label17.Caption = "BALCAO"
Form407.Text1.Text = Form404.Text1.Text
Form407.Label1.Caption = Form404.Label18.Caption
Form404.Adodc2.Refresh
End Sub

Private Sub Command2_Click()
Form404.Text11.Text = 2
If Form2.Label13.Caption <> "Label13" Then
Form406.Label4.Caption = Replace(Form2.Label13.Caption, ",", ".")
End If
  RecompileFrete
  Form407.Label10.Caption = 2
  Form407.Caption = "Resolvendo pagamento do pedido: " & Form404.Label18.Caption
'manter o frete
'entregar
Form407.Show
Form407.Label17.Caption = "ENTREGA"
Form407.Text1.Text = Form404.Text1.Text
Form407.Label1.Caption = Form404.Label18.Caption
Form404.Adodc2.Refresh
End Sub

Public Sub removerfrete()
Dim valorcontab As Single
Dim resp As Integer
Dim numerodopedido  As Integer
resp = MsgBox("Esta opção removerá o valor de Frete !   mesmo assim deseja continuar?", vbYesNo, "Remover o Frete?")
If resp = 6 Then
numerodopedido = Form404.Label18.Caption
Form404.Adodc2.RecordSource = ""

Form404.Adodc2.CommandType = adCmdText

Form404.Adodc2.RecordSource = "SELECT * FROM `at_frete` WHERE `fk_numPedido`='" & numerodopedido & "'"

'Adodc2.Refresh

Form404.Adodc2.Refresh
If Form404.Adodc2.Recordset.BOF = False Then
Form406.Label4 = Replace(Form404.DataGrid2.Columns(1).Value, ",", ".")
Form404.Adodc2.Recordset.Delete
repassevalorsemFrete
Else
valorcontab = Replace(Label4, ".", ",")
Form404.Show

'#todo erro valor conta 0
If Form404.Adodc2.Recordset.BOF = False Then
Form404.Text1.Text = Text2 - valorcontab
End If
Form404.Text1.Text = Format(Form404.Text1.Text, "Currency")

End If


Form404.Adodc2.RecordSource = ""

Form404.Adodc2.CommandType = adCmdText

Form404.Adodc2.RecordSource = "SELECT `frete` FROM `at_frete` WHERE `fk_numPedido` ='" & numerodopedido & "'"

Form404.Adodc2.Refresh
End If

'contabilize
Form407.Text1 = Form404.Text1.Text
End Sub




Public Sub removerfreteSilencio()
Dim valorcontab As Single
Dim resp As Integer
Dim numerodopedido  As Integer
'resp = MsgBox("Esta opção removerá o valor de Frete !   mesmo assim deseja continuar?", vbYesNo, "Remover o Frete?")
'If resp = 6 Then
numerodopedido = Form404.Label18.Caption
Form404.Adodc2.RecordSource = ""

Form404.Adodc2.CommandType = adCmdText

Form404.Adodc2.RecordSource = "SELECT * FROM `at_frete` WHERE `fk_numPedido`='" & numerodopedido & "'"

'Adodc2.Refresh

Form404.Adodc2.Refresh
        If Form404.Adodc2.Recordset.BOF = False Then
        Form406.Label4 = Replace(Form404.DataGrid2.Columns(1).Value, ",", ".")
        Form404.Adodc2.Recordset.Delete
        repassevalorsemFrete
        Else
        valorcontab = Replace(Label4, ".", ",")
        Form404.Show
        If Text2.Text = "Text1" Then
        Text2 = Form407.Text8
        End If
        
        Form404.Text1.Text = Text2 - valorcontab
        Form404.Text1.Text = Format(Form404.Text1.Text, "Currency")
        End If


Form404.Adodc2.RecordSource = ""

Form404.Adodc2.CommandType = adCmdText

Form404.Adodc2.RecordSource = "SELECT `frete` FROM `at_frete` WHERE `fk_numPedido` ='" & numerodopedido & "'"

Form404.Adodc2.Refresh


'contabilize
Form407.Text1 = Form404.Text1.Text
End Sub



Public Sub RecompileFrete()
Dim valordefrete As Double
Dim numerodopedido As Integer
ConServer


Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs.ActiveConnection = con
numerodopedido = Form404.Label18.Caption
Dim valorDosItens As Double
rs.CursorLocation = adUseClient
  sql = "SELECT SUM(`valor`) FROM `at_itens` WHERE `fk_pedido` = '" & numerodopedido & "'"
  rs.Open sql
 ' Form406.Label4 = rs.Fields("valor").Value
  
  If Form404.Adodc2.Recordset.BOF = False Then
  'UPDATE `at_frete` SET `frete` = '3.8' WHERE `at_frete`.`id` = 300;
  rs.Close
         sql = "UPDATE `at_frete` SET `frete` = '" & Form406.Label4 & "' WHERE `fk_numPedido` = '" & numerodopedido & "'"
  rs.Open sql
        recomponhaFrete
         Form404.Text8.Text = Form2.Label13
         
        
   If Form2.Label13.Caption = "Label13" Then
   valordefrete = Form404.DataGrid2.Columns(0).Value
   Else
  valordefrete = Form2.Label13
  End If
  Else
  recomponhaFrete
  End If
  'CommonDialog1.CancelError = True
'trarar erro
On Error GoTo error


 sql = "SELECT SUM(`valor`) FROM `at_itens` WHERE `fk_pedido` = '" & numerodopedido & "'"
  rs.Open sql
  valorDosItens = rs.Fields("SUM(`valor`)").Value
  
  
 If valordefrete <> 0 Then
   Form404.Text1.Text = Format(valorDosItens + Form404.Text8.Text, "currency")
Else
  Form404.Text1.Text = Format(valorDosItens, "currency")
End If
Form407.Text1.Text = Form404.Text1.Text

 Set rs = Nothing
    Exit Sub

error:
'valorDosItens = 0

 'If valordefrete <> 0 Then
  ' Form404.Text1.Text = Format(valorDosItens + Text8.Text, "currency")
'Else
 ' Form404.Text1.Text = Format(valorDosItens, "currency")
'End If
 Set rs = Nothing
Exit Sub
    
    
   
 Set rs = Nothing
Form407.Text1.Text = Form404.Text1

End Sub
Public Sub repassevalorsemFrete()
Dim valordefrete As Double
Dim numerodopedido As Integer
ConServer


Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs.ActiveConnection = con
numerodopedido = Form404.Label18.Caption
Dim valorDosItens As Double
rs.CursorLocation = adUseClient
  sql = "SELECT SUM(`valor`) FROM `at_itens` WHERE `fk_pedido` = '" & numerodopedido & "'"
  rs.Open sql
  
  If Form404.Adodc2.Recordset.BOF = False Then
        
        
         Form404.Text8.Text = 0
        
   
  valordefrete = 0
  End If
  'CommonDialog1.CancelError = True
'trarar erro
On Error GoTo error



  valorDosItens = rs.Fields("SUM(`valor`)").Value
  
  
 If valordefrete <> 0 Then
   Form404.Text1.Text = Format(valorDosItens + Text8.Text, "currency")
Else
  Form404.Text1.Text = Format(valorDosItens, "currency")
End If
 Set rs = Nothing
    Exit Sub

error:
valorDosItens = 0

 If valordefrete <> 0 Then
   Form404.Text1.Text = Format(valorDosItens + Text8.Text, "currency")
Else
  Form404.Text1.Text = Format(valorDosItens, "currency")
End If
 Set rs = Nothing
Exit Sub
    
    
   
 Set rs = Nothing
Form405.Text2.Text = Form404.Text1

End Sub

Public Sub recomponhaFrete()
If Form404.Adodc2.Recordset.BOF = True Then
Dim valordefrete As Double
Dim numerodopedido As Integer
Dim valorDosItens As Double
ConServer


Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset
Label4.Caption = DBLcurre(Label4.Caption)
Set rs.ActiveConnection = con
numerodopedido = Form404.Label18.Caption
 sql = "SELECT * FROM `at_frete` WHERE `fk_numPedido` = '" & numerodopedido & "' ORDER BY `fk_numPedido` DESC"
  rs.Open sql
  If rs.BOF = True Then
  rs.Close
   sql = "INSERT INTO `at_frete` (`id`, `frete`, `fk_numPedido`) VALUES (NULL, '" & Label4.Caption & "', '" & numerodopedido & "')"
  rs.Open sql
'  rs.Close
valordefrete = Replace(Label4.Caption, ".", ",")
Form404.Adodc2.RecordSource = ""

Form404.Adodc2.CommandType = adCmdText

Form404.Adodc2.RecordSource = "SELECT `frete` FROM `at_frete` WHERE `fk_numPedido`='" & numerodopedido & "'"

'Adodc2.Refresh

Form404.Adodc2.Refresh
  Else
  rs.Close
  End If



'sql = "SELECT SUM(`valor`) FROM `at_itens` WHERE `fk_pedido` = '" & numerodopedido & "'"
 ' rs.Open sql
  
  'If Form404.Adodc2.Recordset.BOF = False Then
        
       valordefrete = Label4.Caption
         Text8.Text = Replace(Label4.Caption, ".", ",")
        
   
  'valordefrete = Form401.Text20.Text
  'End If
  'CommonDialog1.CancelError = True
'trarar erro
On Error GoTo error


 sql = "SELECT SUM(`valor`) FROM `at_itens` WHERE `fk_pedido` = '" & numerodopedido & "'"
  rs.Open sql
  valorDosItens = rs.Fields("SUM(`valor`)").Value
  
  
 If valordefrete <> 0 Then
   Form404.Text1.Text = Format(valorDosItens + Text8.Text, "currency")
Else
  Form404.Text1.Text = Format(valorDosItens, "currency")
End If
 Set rs = Nothing
    Exit Sub

error:
 'Text2.Text = DBLcurre(Text2.Text)
valorDosItens = Form407.Text1

 If valordefrete <> 0 Then
   Form404.Text1.Text = Format(valorDosItens + Text8.Text, "currency")
Else
  Form404.Text1.Text = Format(valorDosItens, "currency")
End If

 Set rs = Nothing
Exit Sub
    
    
   
 Set rs = Nothing
Form407.Text1.Text = Form404.Text1
Form405.Text2.Text = Form404.Text1


End If

End Sub




Private Sub Command2_LostFocus()
If Form401.Text30.Text = 1 Then
'ative funcoes de entrega
funcoesdeentrega
ElseIf Form401.Text30.Text = 2 Then
'ative funcoes de busca
funcoesdebusca
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF1 Then
   Call Command2_Click
  End If
    If KeyCode = vbKeyF2 Then
   Call Command1_Click
  End If
 
    
 
End Sub

Public Sub funcoesdeentrega()
  RecompileFrete
  Form407.Label10.Caption = 2
  Form407.Caption = "Resolvendo pagamento do pedido: " & Form404.Label18.Caption
'manter o frete
'entregar
Form407.Show
Form407.Label17.Caption = "ENTREGA"
Form407.Text1.Text = Form404.Text1.Text
Form407.Label1.Caption = Form404.Label18.Caption
Form404.Adodc2.Refresh

End Sub

Public Sub funcoesdebusca()
 Form407.Label10.Caption = 1
If Form2.cmdEntrega.Caption <> "Entrega" Then

 removerfreteSilencio
 End If
'BALCAO

Form407.Visible = True
Form407.Label17.Caption = "BALCAO"
Form407.Text1.Text = Form404.Text1.Text
Form407.Label1.Caption = Form404.Label18.Caption
Form404.Adodc2.Refresh
End Sub

Private Sub Label4_Change()
Label4.Caption = Label4.Caption
End Sub

Private Sub Text1_Change()
If Form406.Text1 <> 5 Then
        If Form401.Text30.Text = 1 Then
        'ative funcoes de entrega
        funcoesdeentrega
        ElseIf Form401.Text30.Text = 2 Then
        'ative funcoes de busca
        funcoesdebusca
        End If
        If Form401.Text30.Text = 0 Then
        Form406.Show
        Else
        Form406.Hide
        End If
End If
End Sub

Private Sub Text2_Change()
Form407.Text8 = Text2

End Sub
