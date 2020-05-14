VERSION 5.00
Begin VB.Form Form405 
   BackColor       =   &H8000000A&
   Caption         =   "Form9"
   ClientHeight    =   5880
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11130
   Icon            =   "pagamentoPedidos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form9"
   Moveable        =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   11130
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   615
      Left            =   5400
      TabIndex        =   19
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox Text10 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   5040
      TabIndex        =   18
      Top             =   2400
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   15
      Text            =   "75,10"
      Top             =   2400
      Width           =   1815
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   8280
      TabIndex        =   13
      Text            =   "Text8"
      Top             =   2640
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   7440
      TabIndex        =   12
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   8280
      TabIndex        =   11
      Text            =   "Text7"
      Top             =   2160
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   8280
      TabIndex        =   10
      Text            =   "Text6"
      Top             =   1560
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   8280
      TabIndex        =   9
      Text            =   "Text5"
      Top             =   960
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   375
      Left            =   480
      TabIndex        =   8
      Text            =   "Text4"
      Top             =   5160
      Width           =   6135
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "pagamentoPedidos.frx":058A
      Left            =   600
      List            =   "pagamentoPedidos.frx":05A9
      TabIndex        =   4
      Text            =   "Forma de pagamento"
      Top             =   240
      Width           =   6375
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   2400
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   3
      Text            =   "75,10"
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   3240
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "F1           Concluir o pagamento"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   0
      Top             =   4080
      Width           =   6135
   End
   Begin VB.Label Label6 
      Caption         =   "Falta"
      Height          =   255
      Left            =   5040
      TabIndex        =   17
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Valor"
      Height          =   255
      Left            =   2760
      TabIndex        =   16
      Top             =   2160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Frete Static"
      Height          =   255
      Left            =   8400
      TabIndex        =   14
      Top             =   480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Troco"
      Height          =   255
      Left            =   2760
      TabIndex        =   7
      Top             =   3000
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   6
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Dinheiro"
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   2160
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "Form405"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Cancelar As Boolean
Dim valor1 As Double
Dim valor2 As Double
Dim valor3 As Double
Dim Revalidar As Boolean
Dim controleTrasito As Boolean
Dim escolhafeita As Boolean

Private Sub Combo1_Click()

Dim Index As Integer
          
            If Combo1.Text <> "Forma de pagamento" Then
            If Combo1.ListIndex = 0 Then
            Text1.Visible = True
            Text3.Visible = True
            Label1.Visible = True
            Label3.Visible = True
            
            
            Else
            Text3.Visible = False
            Text1.Visible = False
            Label1.Visible = False
            Label3.Visible = False
            
            End If
            
            Index = Combo1.ListIndex
            
              Select Case Index
                        Case 0 '(Entregar) Dinheiro
                         If Text3.Text <> "" And Text3.Text <> "R$ 0,00" Then
                          Text4.Text = "TROCO DE  => " & Text3.Text
                          Else
                          Text4.Text = "SEM TROCO"
                          End If
                          Text8.Text = "0"
                           
                          
                        Case 1 '(Entregar) PG Cartão
                         'informçao adicional
                         Text4.Text = "LEVAR MAQUINA DE CARTAO"
                         'chave de index tomada de decisão
                         Text8.Text = "1"
                         'controleTrasito = True
                         formatUniversal
                        
                        Case 2 '(Entregar) Ticket de alimentação
                           'informçao adicional
                               Text1.Visible = True
            Text3.Visible = True
            Label1.Visible = True
            Label3.Visible = True
                         Text4.Text = "PAGAMENTO EM TICKETs"
                         'chave de index tomada de decisão
                         Text8.Text = "2"
                         'controleTrasito = True
                         formatUniversal

                         
                        Case 3 '(Entregar) Pago!
                          'informçao adicional
                         Text4.Text = "PAGO SO ENTREGAR"
                         'chave de index tomada de decisão
                         Text8.Text = "3"
                         'controleTrasito = True
                         formatUniversal
                        
                        
                        Case 4 '(Entregar) Pagará Depois
                             'informçao adicional
                         Text4.Text = "SO ENTREGAR -ANOTAR"
                         'chave de index tomada de decisão
                         Text8.Text = "4"
                         'controleTrasito = True
                         formatUniversal
                         
                         
                         Case 5 '(Balcão) CLIENTE ESTA ESPERANDO
                         'informçao adicional
                         Text4.Text = "CLIENTE NA LOJA ESPERANDO"
                         'chave de index tomada de decisão
                         Text8.Text = "5"
                         'controleTrasito = True
                         formatUniversal
                        
                        
                        Case 6 '(Balcão) Pago! vem busca
                         'informçao adicional
                         Text4.Text = "PAGO CLIENTE VEM BUSCAR"
                         'chave de index tomada de decisão
                         Text8.Text = "6"
                         'controleTrasito = True
                         formatUniversal
                        
                        
                        Case 7 '(Balcão) Pagar na hora que buscar
                         'informçao adicional
                         Text4.Text = "PAGARA NA HORA QUE BUSCAR"
                         'chave de index tomada de decisão
                         Text8.Text = "7"
                         'controleTrasito = True
                         formatUniversal
                        
                        
                        Case 8 '(Balcão) Pagará Depois
                         'informçao adicional
                         Text4.Text = "SO ENTREGAR - ANOTAR "
                         'chave de index tomada de decisão
                         Text8.Text = "8"
                         'controleTrasito = True
                         formatUniversal
                            
                           
                           
                           
                           
                  End Select
 Command2.Enabled = True
Else
 Command2.Enabled = False
End If

If escolhafeita = False Then
        If Index >= 5 Then
        removerfrete
        Else
        recomponhaFrete
        End If
        

Form405.Text2.Text = Form404.Text1
End If
End Sub

Public Sub removerfrete()
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
Form405.Label4 = Form404.DataGrid2.Columns(1).Value
Form404.Adodc2.Recordset.Delete
repassevalorsemFrete
Else
Form404.Text1.Text = Text2 - Replace(Label4, ".", ",")
Form404.Text1.Text = Format(Form404.Text1.Text, "Currency")
End If


Form404.Adodc2.RecordSource = ""

Form404.Adodc2.CommandType = adCmdText

Form404.Adodc2.RecordSource = "SELECT `frete` FROM `at_frete` WHERE `fk_numPedido` ='" & numerodopedido & "'"

Form404.Adodc2.Refresh
End If

'contabilize
Text2.Text = Form404.Text1.Text
End Sub

Private Sub Command1_Click()
If Text1.Text = "" And Text5 <> "" Then
            'DBLcurre (Text2.Text)
            Text1.Text = Text2.Text
            Text5.Text = Text2.Text
            Text6.Text = Text2.Text
            caucular
            End If
valor1 = Replace(Text5, ".", ",")
valor2 = Replace(Text6, ".", ",")
valor3 = Format(valor1 - valor2, "currency")
Text3.Text = Format(valor3, "currency")


Text1.Text = Format(Text1.Text, "currency")
Text2.Text = Format(Text2.Text, "currency")
Text3.Text = Format(Text3.Text, "currency")
  Call Combo1_Click

End Sub

Private Sub Command2_Click()
escolhafeita = True
Cancelar = True
Form501.Show
Form501.Timer1.Interval = 100
End Sub

Private Sub Command2_GotFocus()
If Text1 <> "" And Text1.Visible = True Then
Text1.Text = Replace(Text1.Text, ".", ",")
Text1.Text = Format(Text1.Text, "currency")

 
Text5.Text = Text1
Text6.Text = Text2
caucular
Else

Text5.Text = Text2.Text
Text6.Text = Text2.Text
caucular


End If
End Sub

Private Sub Command3_Click()
Text2.Text = "75,10"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF1 Then
    Call Command2_Click
  End If
End Sub

Private Sub Form_Load()
Form500.Hide
Text1.Text = ""
Text2.Text = ""
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim resp As Integer
'Form2.Hide
'Form405.Hide
'Form403.Hide
If Cancelar = False Then
        resp = MsgBox("Você não mandou imprimir o pedido se você decidir fechar o pedido será cancelado ! Deseja continuar mesmo assim?", vbYesNo, "Cancelar o pedido ?")
        If resp = 6 Then
        
        Form401.Text26.Text = 0
        Else
        Form401.Text26.Text = 1
        End If
Else
Unload Form402
'Unload Form405
Unload Form403
Unload Form403

Form3.Show
'Unload Me
End If
End Sub

Private Sub Text1_Click()
Text1.Text = ""
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
KeyAscii = 0
Text1.Text = ""
End If
If KeyAscii = 13 Then
KeyAscii = 0
            If Text1 <> "" Then
            Text1.Text = Replace(Text1.Text, ".", ",")
            Text1.Text = Format(Text1.Text, "currency")
            
             
            Text5.Text = Text1
            Text6.Text = Text2
            caucular
            Else
            
            Text5.Text = Text2.Text
            Text6.Text = Text2.Text
            caucular
            
            
            End If
End If
End Sub

Private Sub Text1_LostFocus()
If Text1 <> "" Then
Text1.Text = Replace(Text1.Text, ".", ",")
Text1.Text = Format(Text1.Text, "currency")

 
Text5.Text = Text1
Text6.Text = Text2
caucular
Else

Text5.Text = Text2.Text
Text6.Text = Text2.Text
caucular


End If
'valor3 = valor2 - valor1
'Text7.Text = valor3



End Sub

Private Sub Text2_Change()
Text2.Text = Format(Text2.Text, "currency")
Text9.Text = Text2.Text

End Sub

Private Sub Text2_LostFocus()
Text2.Text = Format(Text2.Text, "currency")
End Sub

Private Sub Text3_Change()
Text7.Text = Text3.Text
End Sub

Private Sub Text5_Change()
Text5 = DBLcurre(Text5.Text)
'Text5 = CDbl(Text5.Text)
End Sub

Private Sub Text6_Change()
Text6 = DBLcurre(Text6.Text)
'ext6 = CDbl(Text6.Text)
End Sub

Private Sub Text7_Change()
Text7 = DBLcurre(Text7.Text)

End Sub

Public Sub caucular()
If Text1.Text = "" And Text9 <> "" Then
            'DBLcurre (Text2.Text)
            Text1.Text = Text2.Text
            Text5.Text = Text2.Text
            Text6.Text = Text2.Text
            caucular
            End If
  valor1 = Replace(Text9, ".", ",")

valor2 = Replace(Text6, ".", ",")
valor3 = Format(valor1 - valor2, "currency")
Text3.Text = Format(valor3, "currency")


Text1.Text = Format(Text1.Text, "currency")
Text2.Text = Format(Text2.Text, "currency")
Text3.Text = Format(Text3.Text, "currency")
  Call Combo1_Click


'
'
'valor1 = Replace(Text9, ".", ",")
'valor2 = Replace(Text6, ".", ",")
'valor3 = Format(valor1 - valor2, "currency")
'Text3.Text = Format(valor3, "currency")
'Call Combo1_Click
End Sub

Public Sub formatUniversal()


Text9.Text = Text2.Text
Text6.Text = Text2.Text
Text7.Text = 0
'caucular


'End If
   
End Sub

Private Sub Text8_Change()
If Text1.Visible = True And Text1.Text <> "" Then
Text9.Text = Text1.Text
Text6.Text = Text2.Text
Text7.Text = Text3.Text
End If
End Sub

Public Sub cmdImprimir()
Dim numPedido As Integer
Dim NpTitulo As String
numPedido = Form404.Label18.Caption
If Revalidar = True Then
Printer.Print " -----------------------------------------------"
Printer.Print "MODIFICACAO DO PEDIDO:  ", numPedido
Printer.Print ")"
Printer.Print ""
Printer.Print ""
Printer.Print " -----------------------------------------------"
Printer.Print "PEDIDO MODIFICADO AGUARDE NOVA COMANDA"
Printer.Print " -----------------------------------------------"
Printer.Print "MODIFICACAO  DESTE PEDIDO ", numPedido
Printer.Print " -----------------------------------------------"
Printer.Print " "
Revalidar = False
ConServer

Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs.ActiveConnection = con
rs.CursorLocation = adUseClient



sql = "UPDATE `at_contadorDePedidos` SET `situacaoImpressao` = '1' WHERE `at_contadorDePedidos`.`id` =  '" & numPedido & "'"
                            rs.Open sql
 Set rs = Nothing
End If
comander6
'imprimirCupon
Cancelar = True
Form404.Text9.Text = 1
Unload Form404
Unload Form401
Unload Form402
Form3.Show
End Sub

Public Sub comander6()
ConServer
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



valorFRete = DBLcurre(Form404.DataGrid2.Columns(0).Value)

  
  sql = "INSERT INTO `at_Cupon` (`id`, `nomeEmpresa`, `numPedido`, `nomeCliente`,`endereco`, `telefone`, `referencia`, `loja`, `fk_itens`, `valor_frete`, `obsvacoes`, `total`, `datahora`, `operador`, `valorRecebido`, `valrorPago`, `troco`, `observacoes2`, `formadepagamento`)" & _
  "VALUES (NULL, '" & Form404.Label1.Caption & "', '" & Form404.Label18.Caption & "','" & Form404.Label12.Caption & "','" & Form404.Label13.Caption & "', '" & Form404.Label15.Caption & "', '" & Form404.Label14.Caption & "', '" & Form404.Label17.Caption & "', '" & Form404.Label18.Caption & "','" & valorFRete & "' , '" & Form404.Text2.Text & "', '" & Form404.Text4.Text & "','" & Form404.Label19.Caption & "', '" & Form404.Label20.Caption & "', '" & Form404.Text3.Text & "', '" & Form404.Text4.Text & "', '" & Form404.Text5.Text & "', '" & Form404.Text7.Text & "', '" & Form404.Text6.Text & "')"
  rs.Open sql
  
 

    
    
   
   
   
   
   
   
   
   
   
   
   
   
 Set rs = Nothing


Exit Sub

error:
valorFRete = 0

  
  sql = "INSERT INTO `at_Cupon` (`id`, `nomeEmpresa`, `numPedido`,  `nomeCliente`,`endereco`, `telefone`, `referencia`, `loja`, `fk_itens`, `valor_frete`, `obsvacoes`, `total`, `datahora`, `operador`, `valorRecebido`, `valrorPago`, `troco`, `observacoes2`, `formadepagamento`)" & _
  "VALUES (NULL, '" & Form404.Label1.Caption & "', '" & Form404.Label18.Caption & "','" & Form404.Label12.Caption & "','" & Form404.Label13.Caption & "', '" & Form404.Label15.Caption & "', '" & Form404.Label14.Caption & "', '" & Form404.Label17.Caption & "', '" & Form404.Label18.Caption & "','" & valorFRete & "' , '" & Form404.Text2.Text & "', '" & Form404.Text4.Text & "','" & Form404.Label19.Caption & "', '" & Form404.Label20.Caption & "', '" & Form404.Text3.Text & "', '" & Form404.Text4.Text & "', '" & Form404.Text5.Text & "', '" & Form404.Text7.Text & "', '" & Form404.Text6.Text & "')"
  rs.Open sql

Exit Sub



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
valorDosItens = Text2.Text

 If valordefrete <> 0 Then
   Form404.Text1.Text = Format(valorDosItens + Text8.Text, "currency")
Else
  Form404.Text1.Text = Format(valorDosItens, "currency")
End If
 Set rs = Nothing
Exit Sub
    
    
   
 Set rs = Nothing
Form405.Text2.Text = Form404.Text1


End If

End Sub
