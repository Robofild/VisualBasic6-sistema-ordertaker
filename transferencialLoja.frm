VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.ocx"
Begin VB.Form Form501 
   BorderStyle     =   0  'None
   Caption         =   "Form9"
   ClientHeight    =   3480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8355
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   8355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   360
      Top             =   3120
   End
   Begin VB.Frame Frame1 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8055
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   495
         Left            =   840
         TabIndex        =   1
         Top             =   2040
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   873
         _Version        =   327682
         BorderStyle     =   1
         Appearance      =   1
         Min             =   1e-4
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "lablel3"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   600
         Left            =   1680
         TabIndex        =   3
         Top             =   1200
         Width           =   5055
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Enviando o pedido para a loja:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   7695
      End
   End
End
Attribute VB_Name = "Form501"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Cancelar As Boolean
Dim Revalidar As Boolean
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Sub Form_Load()
Label3.Caption = Form404.Label17
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'Form503.Show
Unload Form500
Unload Form405
Unload Form2
Unload Form503
Form3.Visible = True

End Sub

Private Sub Timer1_Timer()
'Form501.Show
'Form503.Show
Dim indexB As Integer

  For indexB = 1 To 100
  ProgressBar1.Value = indexB / 2
        'Sleep (8)
                                        
  Next indexB
 transferirValorPagamento
  ProgressBar1.Value = 100
'  cmdImprimir
 comander6
  Timer1.Interval = 0
 ' Form503.Show
Unload Form406
Unload Form407

Unload Form501
Unload Form500
Unload Form405
Unload Form401
Form404.Text9.Text = "1"
Unload Form404
Unload Form2
Unload Form503
 Form3.Show
End Sub

Public Sub cmdImprimir()
Dim numPedido As Integer
Dim NpTitulo As String
numPedido = Form404.Label18.Caption
If Revalidar = True Then
'Printer.Print " -----------------------------------------------"
'Printer.Print "MODIFICACAO DO PEDIDO:  ", numPedido
'Printer.Print ")"
'Printer.Print ""
'Printer.Print ""
'Printer.Print " -----------------------------------------------"
'Printer.Print "PEDIDO MODIFICADO AGUARDE NOVA COMANDA"
'Printer.Print " -----------------------------------------------"
'Printer.Print "MODIFICACAO  DESTE PEDIDO ", numPedido
'Printer.Print " -----------------------------------------------"
'Printer.Print " "
Revalidar = False
ConServer

Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs.ActiveConnection = con
rs.CursorLocation = adUseClient



sql = "UPDATE `at_contadorDePedidos` SET `intPedido` = '" & Form404.Label17.Caption & "',`contador` = '1', `situacaoImpressao` = '" & Form407.Label10.Caption & "'  WHERE `at_contadorDePedidos`.`id` =  '" & numPedido & "'"
                            rs.Open sql
 
End If
 Set rs = Nothing
comander6
'imprimirCupon
Cancelar = True
'Form503.Show
Form404.Text9.Text = 1
Unload Form404
Unload Form401
Unload Form402
Unload Form503
Form3.Show

End Sub

Public Sub comander6()
ConServer
Dim numerodopedido As Integer
Dim operador As String
Dim valorFRete As String
Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs.ActiveConnection = con

rs.CursorLocation = adUseClient
'CommonDialog1.CancelError = True
'trarar erro
On Error GoTo error
numerodopedido = Form404.Label18.Caption
Form401.Text26.Text = 1

If Form404.Text10.Text = 1 Then

Form404.Adodc2.RecordSource = ""

Form404.Adodc2.CommandType = adCmdText

Form404.Adodc2.RecordSource = "SELECT `frete` FROM `at_frete` WHERE `fk_numPedido` ='" & numerodopedido & "'"

Form404.Adodc2.Refresh
Form501.ProgressBar1.Value = 60
If Form404.Adodc2.Recordset.BOF = False Then
valorFRete = DBLcurre(Form404.DataGrid2.Columns(0).Value)
Else

sql = "INSERT INTO `at_frete` (`id`, `frete`, `fk_numPedido`) VALUES  (NULL, '0', '" & numerodopedido & "')"
  rs.Open sql
End If



sql = "UPDATE `robofi61_order_taker`.`at_Cupon` SET `nomeCliente` = '" & Form404.Label12.Caption & "', `endereco` ='" & Form404.Label13.Caption & "', `telefone` = '" & Form404.Label15.Caption & "', `referencia` = '" & Form404.Label14.Caption & "', `loja` = '" & Form404.Label17.Caption & "', `valor_frete` ='" & valorFRete & "'  , `obsvacoes` = '" & Form404.Text2.Text & "', `total` = '" & Form407.Label8.Caption & "',`datahora` = '" & Form404.Label19.Caption & "', `operador` = '" & Form404.Label20.Caption & "', `valorRecebido` = '" & Form404.Text3.Text & "' , `valrorPago` = '" & Form404.Text4.Text & "', `troco` = '" & Form404.Text7.Text & "', `observacoes2` = '" & Form2.Text11.Text & "', `formadepagamento` = '" & Form404.Text6.Text & "', `fk_Cliente` = '" & Form2.Text11.Text & "' WHERE (`numPedido` = '" & Form404.Label18.Caption & "')"
  rs.Open sql
  rs.Close
Else
  Form404.Adodc2.RecordSource = ""

Form404.Adodc2.CommandType = adCmdText

Form404.Adodc2.RecordSource = "SELECT `frete` FROM `at_frete` WHERE `fk_numPedido` ='" & numerodopedido & "'"

Form404.Adodc2.Refresh
Form501.ProgressBar1.Value = 60
If Form404.Adodc2.Recordset.BOF = False Then
valorFRete = DBLcurre(Form404.DataGrid2.Columns(0).Value)
Else

sql = "INSERT INTO `at_frete` (`id`, `frete`, `fk_numPedido`) VALUES  (NULL, '0', '" & numerodopedido & "')"
  rs.Open sql
End If

  sql = "INSERT INTO `at_Cupon` (`id`, `nomeEmpresa`, `numPedido`,`nomeCliente`, `endereco`, `telefone`, `referencia`, `loja`, `fk_itens`, `valor_frete`, `obsvacoes`, `total`, `datahora`, `operador`, `valorRecebido`, `valrorPago`, `troco`, `observacoes2`, `formadepagamento`,`fk_Cliente`)" & _
  "VALUES (NULL, '" & Form404.Label1.Caption & "', '" & Form404.Label18.Caption & "','" & Form404.Label12.Caption & "','" & Form404.Label13.Caption & "', '" & Form404.Label15.Caption & "', '" & Form404.Label14.Caption & "', '" & Form404.Label17.Caption & "', '" & Form404.Label18.Caption & "','" & valorFRete & "' , '" & Form404.Text2.Text & "', '" & Form407.Label8.Caption & "','" & Form404.Label19.Caption & "', '" & Form404.Label20.Caption & "', '" & Form404.Text3.Text & "', '" & Form404.Text4.Text & "', '" & Form404.Text5.Text & "', '" & Form404.Text7.Text & "', '" & Form404.Text6.Text & "','" & Form2.Text11.Text & "')"
  rs.Open sql
  
 rs.Close
End If
    
    
   
sql = "UPDATE `at_contadorDePedidos` SET `contador` = '1' WHERE `at_contadorDePedidos`.`id` =  '" & Form404.Label18.Caption & "'"
                            rs.Open sql
   
   
   
   
   
   
   
   
   
   
 Set rs = Nothing


Exit Sub

error:
valorFRete = 0

  
 ' sql = "INSERT INTO `at_Cupon` (`id`, `nomeEmpresa`, `numPedido`, `nomeCliente`, `endereco`, `telefone`, `referencia`, `loja`, `fk_itens`, `valor_frete`, `obsvacoes`, `total`, `datahora`, `operador`, `valorRecebido`, `valrorPago`, `troco`, `observacoes2`, `formadepagamento`,`fk_Cliente`)" & _
 ' "VALUES (NULL, '" & Form404.Label1.Caption & "', '" & Form404.Label18.Caption & "','" & Form404.Label12.Caption & "', '" & Form404.Label13.Caption & "', '" & Form404.Label15.Caption & "', '" & Form404.Label14.Caption & "', '" & Form404.Label17.Caption & "', '" & Form404.Label18.Caption & "','" & valorFRete & "' , '" & Form404.Text2.Text & "', '" & Form404.Text4.Text & "','" & Form404.Label19.Caption & "', '" & Form404.Label20.Caption & "', '" & Form404.Text3.Text & "', '" & Form404.Text4.Text & "', '" & Form404.Text5.Text & "', '" & Form404.Text7.Text & "', '" & Form404.Text6.Text & "','" & Form2.Text11.Text & "')"
 ' rs.Open sql


' rs.Close

    
    
   'SET A IMPRESSAO
sql = "UPDATE `at_contadorDePedidos` SET `intPedido` = '" & Form404.Label17.Caption & "', `contador` = '1',`situacaoImpressao` = '" & Form404.Text11.Text & "' WHERE `at_contadorDePedidos`.`id` =  '" & Form404.Label18.Caption & "'"
                            rs.Open sql


Exit Sub



End Sub



Public Sub transferirValorPagamento()
If Timer1.Interval <> 0 Then
'Form404.Text3.Text = Form405.Text5.Text
'Form404.Text4.Text = Form405.Text6.Text
'Form404.Text5.Text = Form405.Text7.Text
'Form404.Text6.Text = Form405.Text8.Text
'Form404.Text7.Text = Form405.Text4.Text
'Form404.Command2.Enabled = True

Timer1.Interval = 0
End If
End Sub
