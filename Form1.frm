VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Order Taker"
   ClientHeight    =   10395
   ClientLeft      =   2415
   ClientTop       =   2460
   ClientWidth     =   17910
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10395
   ScaleWidth      =   17910
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   17910
      _ExtentX        =   31591
      _ExtentY        =   1111
      ButtonWidth     =   1799
      ButtonHeight    =   953
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   10
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Atendimento"
            Key             =   "at"
            Description     =   "iniciar o atendimento do cliente"
            Object.ToolTipText     =   " Iniciar o atendimento do cliente"
            Object.Tag             =   "Texte"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Cadastro "
            Key             =   "ca"
            Description     =   " cadastro do cliente"
            Object.ToolTipText     =   "Trabalha as funções inerentes ao cadastro  do cliente"
            Object.Tag             =   ""
            ImageIndex      =   13
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Pedido"
            Key             =   "pe"
            Description     =   " prepara o pedido do cliente "
            Object.ToolTipText     =   " Prepara, repara, refaz, pedidos"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Pagamento"
            Key             =   "pa"
            Description     =   "formas de pagamento finalizações"
            Object.ToolTipText     =   "Ações referentes ao pagamento do pedido"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Loja"
            Key             =   "lj"
            Description     =   "distribui para as lojas"
            Object.ToolTipText     =   "Distribui o pedido para  as lojas"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Enviar"
            Key             =   "en"
            Description     =   "envia e finaliza o pedido finaliza o pedido e envia"
            Object.ToolTipText     =   "Finalizar o atendimento"
            Object.Tag             =   "en"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Recebendo"
            Key             =   ""
            Object.Tag             =   ""
            Value           =   1
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Whatsapp"
            Key             =   "wt"
            Object.ToolTipText     =   "Registro de reclamações"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Cardápio"
            Key             =   "cpa"
            Object.ToolTipText     =   "Permite visualizar o conteúdo disponível para pedidos"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Alterações"
            Key             =   "alt"
            Object.ToolTipText     =   "Alterações ou modificações em pedidos já finalizados"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.PictureBox Picture1 
         Height          =   255
         Left            =   6360
         Picture         =   "Form1.frx":0442
         ScaleHeight     =   195
         ScaleWidth      =   555
         TabIndex        =   3
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   11640
         TabIndex        =   1
         Top             =   120
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   661
         _Version        =   327682
         Appearance      =   1
         Max             =   50
      End
   End
   Begin VB.Timer Timer3 
      Interval        =   10000
      Left            =   240
      Top             =   960
   End
   Begin VB.Timer TimerEnvio 
      Left            =   10920
      Top             =   720
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   8160
      TabIndex        =   5
      Top             =   4200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   12360
      TabIndex        =   4
      Top             =   2760
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Timer Timer2 
      Left            =   6840
      Top             =   720
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Top             =   9780
      Width           =   17910
      _ExtentX        =   31591
      _ExtentY        =   1085
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   31547
            MinWidth        =   2152
            Picture         =   "Form1.frx":0E44
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Left            =   12960
      Top             =   1320
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   17040
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   14
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":101E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":11F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":13D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":15AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":1786
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":1960
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":1B3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":1D14
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":1EEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":20C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":22A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":247C
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":2656
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":2830
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnu_arquivo 
      Caption         =   "Arquivo"
      Index           =   0
      Begin VB.Menu mnu_backup 
         Caption         =   "Criar Backup"
         Enabled         =   0   'False
      End
      Begin VB.Menu Mnu_senhaUser 
         Caption         =   "Usuários e Senhas"
      End
      Begin VB.Menu Mnu_logs 
         Caption         =   "Log's (Que usuários Fiseram)"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnu_calculadora 
         Caption         =   "Calculadora"
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu Mnu_Sair 
         Caption         =   "Sair"
         Shortcut        =   ^{F4}
      End
   End
   Begin VB.Menu mnu_cadastro 
      Caption         =   "Cadastros"
      Begin VB.Menu menu_dados 
         Caption         =   "Dados da empresa"
      End
      Begin VB.Menu mnu_cadastrofuncion 
         Caption         =   "Cadastro de Funcionários "
         Enabled         =   0   'False
         WindowList      =   -1  'True
      End
      Begin VB.Menu mnu_frete 
         Caption         =   "Frete"
      End
      Begin VB.Menu mnu_clientes 
         Caption         =   "Clientes"
      End
      Begin VB.Menu mnu_lojas 
         Caption         =   "Lojas (pdv)"
      End
      Begin VB.Menu mnu_cardapio 
         Caption         =   "Cardápio"
      End
   End
   Begin VB.Menu mnu_modulos 
      Caption         =   "Módulos"
      Begin VB.Menu mnu_pedidos 
         Caption         =   "Chamadas Pedidos/ Atendidos"
         Enabled         =   0   'False
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu menu_recibos 
         Caption         =   "Recibos avulsos"
         Enabled         =   0   'False
         Begin VB.Menu mnuEmitir 
            Caption         =   "Emitir recibos"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnu_consultarRecibos 
            Caption         =   "Consultar Recibos"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu mnu_pagamentos 
         Caption         =   "Pagamentos"
         Begin VB.Menu mnu_relatorios 
            Caption         =   "Relatorio de Pagamentos"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu mnu_caixa 
         Caption         =   "consultar movimento"
      End
      Begin VB.Menu line 
         Caption         =   "___________________"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnu_logotipos 
         Caption         =   "Logotipos.."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnu_impressoras 
         Caption         =   "Impressoras..."
      End
      Begin VB.Menu mnu_atualizar 
         Caption         =   "Baixar/Atualizar CEP's.."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnu_servicos 
         Caption         =   "Serviços Google"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnu_suporte 
      Caption         =   "&Suporte Técnico"
      Begin VB.Menu mnu_manual 
         Caption         =   "Manual do Sistema"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnu_videoaulas 
         Caption         =   "Vídeo-aulas na internet"
         Enabled         =   0   'False
      End
      Begin VB.Menu line2 
         Caption         =   "___________________"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnu_supremoto 
         Caption         =   "SuporteRemoto TeamViewer"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuRegistrodoSoftawares 
         Caption         =   "Software registrado para TELE BH AMARELINHO"
      End
   End
   Begin VB.Menu nmu_atalho 
      Caption         =   "Atalho"
      Begin VB.Menu Mnu_atendimento 
         Caption         =   "Atendimento"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub Command1_Click()

Dim idPiloto As Integer
Dim ContadorPiloto As Integer
Dim PedidoEntrege As Integer
'obter parra o cupom
Dim nomeEmpresa  As String
Dim nomeDoCliente As String
Dim numeroPedido As Integer
Dim endereco As String
Dim telefone As String
Dim referenciaEntrega As String
Dim loja As String
Dim valorFRete As Double
Dim observacoesDopedido As String
Dim TotalDaCompra As Double
Dim Datahora As String
Dim operador As String
Dim ValorRecebido As Double
Dim ValorPago As Double
Dim Troco As Double
Dim ObservacoesparaEntrega As String
Dim formaPagamento As Integer

'itens
Dim QtditensPorLista As Integer
Dim indexB As Integer
Dim QtdP As Integer
Dim descricaoP As String
Dim PrecoP As Double

'fonte padrao
Dim fontepadrao As Integer
Dim OldFont
indexB = 0

                                          
                                          
                                          
                                          
ConServer


Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs.ActiveConnection = con

rs.CursorLocation = adUseClient
'trasendocontador de pedidos ultimo com contador 1

  sql = "SELECT * FROM `at_contadorDePedidos` WHERE `contador` = 1 ORDER BY `id` DESC LIMIT 1"
  rs.Open sql
  idPiloto = rs.Fields("id").Value
  rs.Close

          ' trazer cupom combinado ao ultimo contador
            sql = "SELECT * FROM `at_Cupon` WHERE `numPedido` = '" & idPiloto & "'ORDER BY `numPedido` DESC LIMIT 1"
            rs.Open sql
             
            'Recebendo dados para o cupom
                nomeEmpresa = rs.Fields("nomeEmpresa").Value
                'nomeDoCliente = rs.Fields("nomeCliente").Value
                numeroPedido = rs.Fields("numPedido").Value
                endereco = rs.Fields("endereco").Value
                telefone = rs.Fields("telefone").Value
                referenciaEntrega = rs.Fields("referencia").Value
                loja = rs.Fields("loja").Value
                valorFRete = rs.Fields("valor_frete").Value
                observacoesDopedido = rs.Fields("obsvacoes").Value
                TotalDaCompra = rs.Fields("total").Value
                Datahora = rs.Fields("datahora").Value
                operador = rs.Fields("operador").Value
                ValorRecebido = rs.Fields("valorRecebido").Value
                ValorPago = rs.Fields("valrorPago").Value
                Troco = rs.Fields("troco").Value
                ObservacoesparaEntrega = rs.Fields("observacoes2").Value
                formaPagamento = rs.Fields("formadepagamento").Value
            
            
            
            
            rs.Close
                          
                          
                          
                          
                          'obter a quantidade de loopings melhor dizer ober quantidade de itens
                            sql = "SELECT COUNT(*) FROM `at_itens` WHERE `fk_pedido` = '" & idPiloto & "'"
                            rs.Open sql
                            QtditensPorLista = rs.Fields("COUNT(*)").Value
                            rs.Close
                                        
   'itens do ultimo cupom combinidos com o contador
                                        
                                          sql = "SELECT * FROM `at_itens` WHERE `fk_pedido` ='" & idPiloto & "'"
                                          rs.Open sql
                                          nomeDoCliente = rs.Fields("valor").Value
                                          'For indexB = 20 To 1 Step -3
                                          'sb.Append (indexB.ToString)
                                          'sb.Append (" ")
                                          'Next indexB
                                        
                                           
' inicio da impressão dinheiro entrega
'fontepadrao = Printer.FontSize
' Dim PrinterFont As New StdFont
 
OldFont = Printer.FontName   ' Preserva a fonte original.
Printer.Print Tab(10); Format(nomeEmpresa, ">")
Printer.Print
Printer.Print " -----------------------------------------------"

Printer.FontSize = "15"
If formaPagamento <= 4 Then
Printer.Print " ( ENTREGA )       Nº PEDIDO:  "; numeroPedido
Else
        If formaPagamento = 10 Then
        Printer.Print " *******CANCELAR+++++++!!!!!!!!!       Nº PEDIDO:  "; numeroPedido
        Else
        Printer.Print " ( BALCAO )       Nº PEDIDO:  "; numeroPedido
        End If
End If
Printer.FontSize = "10"
Printer.Print " -----------------------------------------------"
Printer.Print "TELEFONE:  "; telefone, nomeDoCliente
Printer.Print "ENDERECO:   "; Format(endereco, ">")
Printer.Print ""
Printer.Print "REFERENCIA: ", Format(referenciaEntrega, ">")
Printer.Print " -----------------------------------------------"
Printer.Print "LOJA:  "; Format(loja, ">")
Printer.Print " -----------------------------------------------"
Printer.Print "QUANTIDADE / DESCRICAO               / PRECO "
Printer.Print " -----------------------------------------------"
Printer.Print " "
                                          For indexB = 1 To QtditensPorLista
                                          QtdP = rs.Fields("Quantidade").Value
                                          descricaoP = rs.Fields("descrição").Value
                                          PrecoP = rs.Fields("valor").Value
                                          Printer.Print " -----------------------------------------------"
                                           If formaPagamento = 10 Then
        Printer.Print " *******  CANCELAR  +++++++!!!!!!!!!       Nº PEDIDO:  "; numeroPedido
        End If
                                          
                                          Printer.Print "  "; QtdP, descricaoP, Spc(10); Format(PrecoP, "Currency")
                                          Printer.Print
                                          If QtditensPorLista = 1 Then
                                            Printer.Print
                                            Printer.Print
                                            Printer.Print
                                            Printer.Print
                                          End If
                                          
                                          Printer.Print
                                          Printer.Print " -----------------------------------------------"
                                          rs.MoveNext
                                          Next indexB

Printer.Print " OBS  "; observacoesDopedido
Printer.Print " -----------------------------------------------"
Printer.Print Tab(15); " SOMATORIA"
Printer.Print
Printer.Print " -----------------------------------------------"
        Select Case formaPagamento
            Case 0 '(Entregar) Dinheiro
                                          Printer.Print "DINHEIRO.............", ; Format(ValorRecebido, "Currency")
                                          Printer.Print
                                          Printer.Print "TAXA DE ENTREGA -> ", ; Format(valorFRete, "Currency")
                                          Printer.Print
                                          Printer.Print "TOTAL...................", ; Format(TotalDaCompra, "Currency")
                                          Printer.Print
                                          Printer.Print
                                          Printer.Print
                                          Printer.Print "TROCO...................", ; Format(Troco, "Currency")
 
            Case 1 '(Entregar) PG Cartão
                                          Printer.Print "CARTAO"
                                          Printer.Print
                                          Printer.Print "TAXA DE ENTREGA -> ", ; Format(valorFRete, "Currency")
                                          Printer.Print
                                          Printer.Print "TOTAL...................", ; Format(TotalDaCompra, "Currency")

                                          
            Case 2 '(Entregar) Ticket de alimentação
                                          Printer.Print "TICKET"
                                          Printer.Print
                                          Printer.Print "TAXA DE ENTREGA -> ", ; Format(valorFRete, "Currency")
                                          Printer.Print
                                          Printer.Print "TOTAL...................", ; Format(TotalDaCompra, "Currency")

            Case 3 '(Entregar) Pago!
                                          Printer.Print "PAGO"
                                          Printer.Print
                                          Printer.Print "TAXA DE ENTREGA -> ", ; Format(valorFRete, "Currency")
                                          Printer.Print
                                          Printer.Print "TOTAL...................", ; Format(TotalDaCompra, "Currency")


            
            Case 4 ' (Entregar) Pagará depois-Anotar
                                          Printer.Print "ANOTAR"
                                          Printer.Print
                                          Printer.Print "TAXA DE ENTREGA -> ", ; Format(valorFRete, "Currency")
                                          Printer.Print
                                          Printer.Print "TOTAL...................", ; Format(TotalDaCompra, "Currency")

            Case 5 '(Balcão) Cliente aguarda na Loja
                                          Printer.Print "CLIENTE NA LOJA"
                                          Printer.Print
                                          Printer.Print "TAXA DE ENTREGA -> ", ; Format(valorFRete, "Currency")
                                          Printer.Print
                                          Printer.Print "TOTAL...................", ; Format(TotalDaCompra, "Currency")

            Case 6 ' (Balcão) Pago! vem buscar
                                          Printer.Print "VEM BUSCAR PAGO!"
                                          Printer.Print
                                          Printer.Print "TAXA DE ENTREGA -> ", ; Format(valorFRete, "Currency")
                                          Printer.Print
                                          Printer.Print "TOTAL...................", ; Format(TotalDaCompra, "Currency")
            Case 7 ' (Balcão) Pago! vem buscar
                                          Printer.Print "VEM BUSCAR PAGO!"
                                          Printer.Print
                                          Printer.Print "TAXA DE ENTREGA -> ", ; Format(valorFRete, "Currency")
                                          Printer.Print
                                          Printer.Print "TOTAL...................", ; Format(TotalDaCompra, "Currency")
            Case 8 ' (Balcão) Pagar na hora que buscar
                                          Printer.Print "VEM BUSCAR / RECEBER!"
                                          Printer.Print
                                          Printer.Print "TAXA DE ENTREGA -> ", ; Format(valorFRete, "Currency")
                                          Printer.Print
                                          Printer.Print "TOTAL...................", ; Format(TotalDaCompra, "Currency")
            Case 9 ' (Anotar) Pagará Depois
                                          Printer.Print "ANOTAR"
                                          Printer.Print
                                          Printer.Print "TAXA DE ENTREGA -> ", ; Format(valorFRete, "Currency")
                                          Printer.Print
                                          Printer.Print "TOTAL...................", ; Format(TotalDaCompra, "Currency")
            Case 10 ' (Cancelar) Cancelar
                                          Printer.Print "*******CANCELAR "
                                          Printer.Print
                                          Printer.Print "NÃO FAZER NAO ENTREGAR  -> ", ; Format(valorFRete, "Currency")
                                          Printer.Print
                                          Printer.Print "***CANCELADO***.", ; Format(TotalDaCompra, "Currency")

        End Select


Printer.Print " ";
Printer.Print "";
Printer.Print " ";
Printer.Print " -----------------------------------------------"
Printer.Print , ObservacoesparaEntrega
Printer.Print " OP > "; operador; Spc(2); "REG"; Datahora


Printer.EndDoc
                                          rs.Close
                                        

                            
                        
                        
                        
                        
                  






 Set rs = Nothing



End Sub

Private Sub Command2_Click()
'imprimirCupon
End Sub

Public Sub imprimirCupon(numPedido As Integer)

Dim idPiloto As Integer
Dim ContadorPiloto As Integer
Dim PedidoEntrege As Integer
'obter parra o cupom
Dim nomeEmpresa  As String
Dim nomeDoCliente As String
Dim numeroPedido As Integer
Dim endereco As String
Dim telefone As String
Dim referenciaEntrega As String
Dim loja As String
Dim valorFRete As Double
Dim observacoesDopedido As String
Dim TotalDaCompra As Double
Dim Datahora As String
Dim operador As String
Dim ValorRecebido As Double
Dim ValorPago As Double
Dim Troco As Double
Dim ObservacoesparaEntrega As String
Dim formaPagamento As Integer

'itens
Dim QtditensPorLista As Integer
Dim indexB As Integer
Dim QtdP As Integer
Dim descricaoP As String
Dim PrecoP As Double

'fonte padrao
Dim fontepadrao As Integer
Dim OldFont
indexB = 0

                                          
                                          
                                          
                                          
ConServer


Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs.ActiveConnection = con

rs.CursorLocation = adUseClient
'trasendocontador de pedidos ultimo com contador 1
' rs.Close
  'sql = "SELECT * FROM `at_contadorDePedidos` WHERE `contador` = 1 ORDER BY `id` DESC LIMIT 1"
  'rs.Open sql
  idPiloto = numPedido
  'rs.Close

          ' trazer cupom combinado ao ultimo contador
            sql = "SELECT * FROM `at_Cupon` WHERE `numPedido` = '" & idPiloto & "'ORDER BY `id` DESC LIMIT 1"
            rs.Open sql
If rs.BOF = False Then
            'Recebendo dados para o cupom
                nomeEmpresa = rs.Fields("nomeEmpresa").Value
                'nomeDoCliente = rs.Fields("nomeCliente").Value
                numeroPedido = rs.Fields("numPedido").Value
                endereco = rs.Fields("endereco").Value
                telefone = rs.Fields("telefone").Value
                referenciaEntrega = rs.Fields("referencia").Value
                loja = rs.Fields("loja").Value
                valorFRete = Format(rs.Fields("valor_frete").Value, "General Number")
                observacoesDopedido = rs.Fields("obsvacoes").Value
                TotalDaCompra = rs.Fields("total").Value
                Datahora = rs.Fields("datahora").Value
                operador = rs.Fields("operador").Value
                ValorRecebido = rs.Fields("valorRecebido").Value
                ValorPago = rs.Fields("valrorPago").Value
                Troco = rs.Fields("troco").Value
                ObservacoesparaEntrega = rs.Fields("observacoes2").Value
                formaPagamento = rs.Fields("formadepagamento").Value
            
            
            
            
            rs.Close
                          
                          
                          
                          
                          'obter a quantidade de loopings melhor dizer ober quantidade de itens
                            sql = "SELECT COUNT(*) FROM `at_itens` WHERE `fk_pedido` = '" & idPiloto & "'"
                            rs.Open sql
                            If rs.BOF = False Then
                            QtditensPorLista = rs.Fields("COUNT(*)").Value
                            Else
                            QtditensPorLista = 0
                            formaPagamento = 10
                            End If
                            rs.Close
                                        
   'itens do ultimo cupom combinidos com o contador
                                        
                                          sql = "SELECT * FROM `at_itens` WHERE `fk_pedido` ='" & idPiloto & "'"
                                          rs.Open sql
                                          If rs.BOF = False Then
                                          nomeDoCliente = rs.Fields("fk_cliente").Value
                                          Else
                                          nomeDoCliente = "CANCELADO"
                                          End If
                                          'For indexB = 20 To 1 Step -3
                                          'sb.Append (indexB.ToString)
                                          'sb.Append (" ")
                                          'Next indexB
                                        
                                           
' inicio da impressão dinheiro entrega
'fontepadrao = Printer.FontSize
' Dim PrinterFont As New StdFont
 
OldFont = Printer.FontName   ' Preserva a fonte original.
Printer.Print Tab(10); Format(nomeEmpresa, ">")
Printer.Print
Printer.Print " -----------------------------------------------"

Printer.FontSize = "15"
If formaPagamento <= 4 Then
Printer.Print " ( ENTREGA )       Nº PEDIDO:  "; numeroPedido
Else
        If formaPagamento = 10 Then
        Printer.Print " *******CANCELAR+++++++!!!!!!!!!       Nº PEDIDO:  "; numeroPedido
        Else
        Printer.Print " ( BALCAO )       Nº PEDIDO:  "; numeroPedido
        End If
End If
Printer.FontSize = "10"
Printer.Print " -----------------------------------------------"
Printer.Print "TELEFONE:  "; telefone, nomeDoCliente
Printer.Print "ENDERECO:   "; Format(endereco, ">")
Printer.Print ""
Printer.Print "REFERENCIA: ", Format(referenciaEntrega, ">")
Printer.Print " -----------------------------------------------"
Printer.Print "LOJA:  "; Format(loja, ">")
Printer.Print " -----------------------------------------------"
Printer.Print "QUANTIDADE / DESCRICAO               / PRECO "
Printer.Print " -----------------------------------------------"
Printer.Print " "
                                          For indexB = 1 To QtditensPorLista
                                          QtdP = rs.Fields("Quantidade").Value
                                          descricaoP = rs.Fields("descrição").Value
                                          PrecoP = rs.Fields("valor").Value
                                          Printer.Print " -----------------------------------------------"
                                           If formaPagamento = 10 Then
        Printer.Print " *******  CANCELAR  +++++++!!!!!!!!!       Nº PEDIDO:  "; numeroPedido
        End If
                                          
                                          Printer.Print "  "; QtdP, descricaoP
                                          Printer.Print "-----"; Format(PrecoP, "Currency")
                                          If QtditensPorLista = 1 Then
                                            Printer.Print
                                            Printer.Print
                                            Printer.Print
                                            Printer.Print
                                          End If
                                          
                                          Printer.Print
                                          Printer.Print " -----------------------------------------------"
                                          rs.MoveNext
                                          Next indexB

Printer.Print " OBS  "; observacoesDopedido
Printer.Print " -----------------------------------------------"
Printer.Print Tab(15); " SOMATORIA"
Printer.Print
Printer.Print " -----------------------------------------------"
        Select Case formaPagamento
            Case 0 '(Entregar) Dinheiro
                                          Printer.Print "DINHEIRO.............", ; Format(ValorRecebido, "Currency")
                                          Printer.Print
                                          Printer.Print "TAXA DE ENTREGA -> ", ; Format(valorFRete, "Currency")
                                          Printer.Print
                                          Printer.Print "TOTAL...................", ; Format(TotalDaCompra, "Currency")
                                          Printer.Print
                                          Printer.Print
                                          Printer.Print
                                          Printer.Print "TROCO...................", ; Format(Troco, "Currency")
 
            Case 1 '(Entregar) PG Cartão
                                          Printer.Print "CARTAO"
                                          Printer.Print
                                          Printer.Print "TAXA DE ENTREGA -> ", ; Format(valorFRete, "Currency")
                                          Printer.Print
                                          Printer.Print "TOTAL...................", ; Format(TotalDaCompra, "Currency")

                                          
            Case 2 '(Entregar) Ticket de alimentação
                                          Printer.Print "TICKET"
                                          Printer.Print
                                          Printer.Print "TAXA DE ENTREGA -> ", ; Format(valorFRete, "Currency")
                                          Printer.Print
                                          Printer.Print "TOTAL...................", ; Format(TotalDaCompra, "Currency")

            Case 3 '(Entregar) Pago!
                                          Printer.Print "PAGO"
                                          Printer.Print
                                          Printer.Print "TAXA DE ENTREGA -> ", ; Format(valorFRete, "Currency")
                                          Printer.Print
                                          Printer.Print "TOTAL...................", ; Format(TotalDaCompra, "Currency")


            
            Case 4 ' (Entregar) Pagará depois-Anotar
                                          Printer.Print "ANOTAR"
                                          Printer.Print
                                          Printer.Print "TAXA DE ENTREGA -> ", ; Format(valorFRete, "Currency")
                                          Printer.Print
                                          Printer.Print "TOTAL...................", ; Format(TotalDaCompra, "Currency")

            Case 5 '(Balcão) Cliente aguarda na Loja
                                          Printer.Print "CLIENTE NA LOJA"
                                          Printer.Print
                                          Printer.Print "TAXA DE ENTREGA -> ", ; Format(valorFRete, "Currency")
                                          Printer.Print
                                          Printer.Print "TOTAL...................", ; Format(TotalDaCompra, "Currency")

            Case 6 ' (Balcão) Pago! vem buscar
                                          Printer.Print "VEM BUSCAR PAGO!"
                                          Printer.Print
                                          Printer.Print "TAXA DE ENTREGA -> ", ; Format(valorFRete, "Currency")
                                          Printer.Print
                                          Printer.Print "TOTAL...................", ; Format(TotalDaCompra, "Currency")
            Case 7 ' (Balcão) Pago! vem buscar
                                          Printer.Print "VEM BUSCAR PAGO!"
                                          Printer.Print
                                          Printer.Print "TAXA DE ENTREGA -> ", ; Format(valorFRete, "Currency")
                                          Printer.Print
                                          Printer.Print "TOTAL...................", ; Format(TotalDaCompra, "Currency")
            Case 8 ' (Balcão) Pagar na hora que buscar
                                          Printer.Print "VEM BUSCAR / RECEBER!"
                                          Printer.Print
                                          Printer.Print "TAXA DE ENTREGA -> ", ; Format(valorFRete, "Currency")
                                          Printer.Print
                                          Printer.Print "TOTAL...................", ; Format(TotalDaCompra, "Currency")
            Case 9 ' (Anotar) Pagará Depois
                                          Printer.Print "ANOTAR"
                                          Printer.Print
                                          Printer.Print "TAXA DE ENTREGA -> ", ; Format(valorFRete, "Currency")
                                          Printer.Print
                                          Printer.Print "TOTAL...................", ; Format(TotalDaCompra, "Currency")
            Case 10 ' (Cancelar) Cancelar
                                          Printer.Print "*******CANCELAR "
                                          Printer.Print
                                          Printer.Print "NÃO FAZER NAO ENTREGAR  -> ", ; Format(valorFRete, "Currency")
                                          Printer.Print
                                          Printer.Print "***CANCELADO***.", ; Format(TotalDaCompra, "Currency")

        End Select


Printer.Print " ";
Printer.Print "";
Printer.Print " ";
Printer.Print " -----------------------------------------------"
Printer.Print , ObservacoesparaEntrega
Printer.Print " OP > "; operador; Spc(2); "REG"; Datahora
Printer.Print " teste"
Printer.Print ; " ½"


Printer.EndDoc
                                          rs.Close
                               'verifique o pedido
                              'pedido impresso
                             sql = "UPDATE `at_contadorDePedidos` SET `situacaoImpressao` = '0' WHERE `at_contadorDePedidos`.`id` =  '" & idPiloto & "'"
                            rs.Open sql
                                          
                                        

                            
                        
                        
                        
                        
                  



Else
rs.Close
End If


 Set rs = Nothing



End Sub

Private Sub Command3_Click()
'checar contador
ConServer
Dim numeroPeditoFiltrado As Integer
Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs.ActiveConnection = con

rs.CursorLocation = adUseClient

 'apanhar numero do  PEDIDO OFICIAL
        
         sql = "SELECT * FROM `at_contadorDePedidos` WHERE `intPedido` LIKE '%GLORIA%' AND `situacaoImpressao` = 1 ORDER BY `id`  DESC LIMIT 1"
         rs.Open sql
         numeroPeditoFiltrado = rs.Fields("id").Value
         If rs.BOF = False Then

imprimirCupon (numeroPeditoFiltrado)
End If
 Set rs = Nothing

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

End
End Sub

Private Sub menu_dados_Click()
Form12.Show
End Sub

Private Sub Mnu_atendimento_Click()
Form500.Visible = True
Unload Form3

 Form3.Show
End Sub

Private Sub mnu_caixa_Click()
Form10.Show
End Sub

Private Sub mnu_calculadora_Click()
Dim command As String
'command = "c:\windows\notepad.exe"
command = "start calc"
Shell "cmd.exe /c " & command
End Sub

Private Sub mnu_cardapio_Click()
Form50.Show
End Sub

Private Sub mnu_frete_Click()
Form24.Show
'Text1.Text = "1"

End Sub

Private Sub mnu_impressoras_Click()
Form13.Show

End Sub

Private Sub Mnu_logs_Click()
Form10.Show
End Sub

Private Sub mnu_lojas_Click()
Form13.Show
End Sub

Private Sub Mnu_Sair_Click()
Dialog.Show
Form1.StatusBar1.Panels.Remove (2)

End Sub

Private Sub Mnu_senhaUser_Click()
Form8.Show


End Sub

Private Sub Timer1_Timer()
ProgressBar1.Value = ProgressBar1.Value + 1
If ProgressBar1.Value >= 49 Then
Timer1.Interval = 0
ProgressBar1.Value = 0
Form1.Toolbar1.Buttons(1).Value = tbrUnpressed
Form1.Toolbar1.Buttons(2).Value = tbrUnpressed
Form1.Toolbar1.Buttons(3).Value = tbrUnpressed
Form1.Toolbar1.Buttons(4).Value = tbrUnpressed
Form1.Toolbar1.Buttons(5).Value = tbrUnpressed
Form1.Toolbar1.Buttons(6).Value = tbrUnpressed
Form1.Toolbar1.Buttons(7).Value = tbrUnpressed
MsgBox "pedido enviado com sucesso", , "Sucesso!"
Unload Form6
Unload Form2

End If
End Sub

Private Sub Timer3_Timer()
'trarar erro
On Error GoTo error

Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset
Set rs.ActiveConnection = con

   
    sql = "SELECT * FROM `Conexao` WHERE `id` = 1"
    rs.Open sql
    Set rs = Nothing
Exit Sub

error:
ConServer
ConServerloc
Exit Sub


    
    
    
    
End Sub

Private Sub TimerEnvio_Timer()
'checar contador
ConServer
Dim numeroPeditoFiltrado As Integer
Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs.ActiveConnection = con

rs.CursorLocation = adUseClient

 'apanhar numero do  PEDIDO OFICIAL
        
         sql = "SELECT * FROM `at_contadorDePedidos` WHERE `intPedido` LIKE '%GLORIA%' AND `situacaoImpressao` = 1 ORDER BY `id`  DESC LIMIT 1"
         rs.Open sql
         numeroPeditoFiltrado = rs.Fields("id").Value
         If rs.BOF = False Then

imprimirCupon (numeroPeditoFiltrado)
End If
 Set rs = Nothing

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
If (Toolbar1.Buttons(1).Value = tbrPressed) Then
'Unload (Form2)
End If

Select Case Button.Key
Case "at"
 ' Atendimento
 Form500.Visible = True
 Form3.Show
 'Toolbar1.Buttons(1).Value = tbrPressed
 
    
Case "ca"
'cadastro
Form22.Show
Case "pe"
 ' pedido
 Form10.Show
Case "pa"
 ' pagamento
 Form10.Show
Case "lj"
 ' Loja
 Form25.Show
Case "en"

 'finalizar enviar
Case "wt"
Dim command As String
 'whatsapp
command = "C:\Order_Taker\whatsapp.bat"
Shell "cmd.exe /c " & command
 Case "cpa"
 ' cardapio
 Form52.Show
 Form52.Text2.Text = 1
 Case "alt"
 ' Alterações
 Form3.Show
 
 
End Select




End Sub

