VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Form10 
   Caption         =   "Registro de Movimentos"
   ClientHeight    =   8670
   ClientLeft      =   2130
   ClientTop       =   3045
   ClientWidth     =   17310
   Icon            =   "registrodemovimentouser.frx":0000
   LinkTopic       =   "Form10"
   ScaleHeight     =   8670
   ScaleWidth      =   17310
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   11880
      TabIndex        =   13
      Text            =   "Combo1"
      Top             =   1200
      Width           =   2415
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   9240
      TabIndex        =   12
      Text            =   "Combo1"
      Top             =   1200
      Width           =   2415
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3720
      TabIndex        =   11
      Text            =   "Combo1"
      Top             =   1200
      Width           =   2415
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2370
      Left            =   6480
      TabIndex        =   7
      Top             =   1200
      Visible         =   0   'False
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   100728833
      CurrentDate     =   43645
   End
   Begin VB.Frame Frame1 
      Caption         =   "Consultar"
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   16695
      Begin VB.CheckBox Check6 
         Caption         =   "Check3"
         Enabled         =   0   'False
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   11640
         TabIndex        =   18
         Top             =   240
         Width           =   255
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Check3"
         Enabled         =   0   'False
         Height          =   195
         Left            =   9000
         TabIndex        =   17
         Top             =   240
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Check3"
         Enabled         =   0   'False
         Height          =   195
         Left            =   6240
         TabIndex        =   16
         Top             =   240
         Width           =   255
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Check3"
         Enabled         =   0   'False
         Height          =   195
         Left            =   3480
         TabIndex        =   15
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Limpar"
         Height          =   855
         Left            =   14400
         Picture         =   "registrodemovimentouser.frx":0A02
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdPedidos 
         Caption         =   "Pedidos"
         Height          =   855
         Left            =   9000
         Picture         =   "registrodemovimentouser.frx":1404
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   2415
      End
      Begin VB.CommandButton CmdConsultar 
         Caption         =   "Consultar"
         Height          =   855
         Left            =   15600
         Picture         =   "registrodemovimentouser.frx":1E06
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Imprimir"
         Height          =   855
         Left            =   11640
         Picture         =   "registrodemovimentouser.frx":2808
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   2415
      End
      Begin VB.Frame Frame2 
         Caption         =   "Operadores"
         Height          =   975
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   3015
         Begin VB.CheckBox Check2 
            Caption         =   "Saidas do sistma"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   600
            Width           =   2175
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Entradas no sistema"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.CommandButton CmdNOme 
         Caption         =   "Cliente"
         Height          =   855
         Left            =   3480
         Picture         =   "registrodemovimentouser.frx":320A
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   2415
      End
      Begin VB.CommandButton cmdData 
         Caption         =   "Data"
         Height          =   855
         Left            =   6240
         Picture         =   "registrodemovimentouser.frx":3C0C
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   2415
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   6855
      Left            =   240
      TabIndex        =   10
      Top             =   1560
      Width           =   16695
      _ExtentX        =   29448
      _ExtentY        =   12091
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Resultado de Pesquisas "
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'consultar cliente
Dim consultar As Integer
Dim InformacaoDoCliente As String
Dim consultaLocal As String
Dim indexCliente As Integer
'pedidos

Dim consultarpedidos As Integer
Dim InformacaoDopedidos As String
Dim consultaLocalpedidos As String
Dim pedidos As Integer
Dim descProdutos As String

Dim IdCliente As Integer
Dim numPedido As Integer
Dim DataPedido As String
Dim NomeCliente As String
Dim Telefonecliente As String
Dim BairroCliente As String

Dim Operadores As String

'contador de passos
Dim botao As Integer


Dim dataconsulta As String





Private Sub CmdConsultar_Click()
Select Case consultar
'nome entrada e saida 21
'nome saida 5
            Case 18
            'data entrada e saida
            Case 7
            'data entrada no sistema
            Case 10
            'nome entrada
            Case 2
            Case 22
            'data nome entrada e saida
            Case 4
    End Select
    Check2.Value = 0
    Check1.Value = 0
    consultar = 0
End Sub


Private Sub cmdData_Click()
botao = 2


MonthView1.Visible = True

   
 

End Sub

Private Sub CmdNOme_Click()
Dim ValorDefalt As String
botao = 1
indexCliente = Combo1.ListIndex

ValorDefalt = FuncaoRetornevalorDefalt(indexCliente)

If Combo1.Text <> "Opções..." Then
   If Check3.Value = 1 Then
    consultaLocal = InputBox("Digite " & Combo1.Text & " para iniciar a consulta! ", "Consultar " & Combo1.Text, ValorDefalt)
   Else
   consultaLocal = InputBox("Digite " & Combo1.Text & " para iniciar a consulta! ", "Consultar " & Combo1.Text)
   End If
   
   If consultaLocal <> "" Then
   
    InformacaoDoCliente = ("%" + consultaLocal + "%")
   End If
  
End If

 If Combo1.Text <> "Opções..." Then
    IniciarconsultaCliente (indexCliente)
   
 End If


End Sub

Private Sub CmdNOme_GotFocus()
MonthView1.Visible = False
End Sub

Private Sub cmdPedidos_Click()
'Dim consultarpedidos As Integer
'Dim InformacaoDopedidos As String
'Dim consultaLocalpedidos As String
'Dim pedidos As Integer


MonthView1.Visible = False

botao = 3
pedidos = Combo2.ListIndex
Dim ValorDefalt As String
ValorDefalt = FuncaoRetornevalorDefaltcombo2(pedidos)

If Combo2.Text <> "Opções..." Then
   If Check4.Value = 1 Then
    consultaLocalpedidos = InputBox("Digite " & Combo2.Text & " para iniciar a consulta! ", "Consultar " & Combo2.Text, Trim(ValorDefalt))
   Else
   consultaLocalpedidos = InputBox("Digite " & Combo2.Text & " para iniciar a consulta! ", "Consultar " & Combo2.Text)
   End If



   
   If consultaLocalpedidos <> "" Then
   
    InformacaoDoCliente = ("%" + consultaLocalpedidos + "%")
   End If
  
End If

 If Combo2.Text <> "Opções..." Then
    IniciarconsultaPedido (pedidos)
   
 End If
End Sub

Private Sub Combo1_Click()
Call CmdNOme_Click
End Sub
Private Sub Combo2_Click()
Call cmdPedidos_Click
End Sub
Private Sub Command1_Click()
MonthView1.Visible = False
End Sub

Private Sub Command1_GotFocus()
MonthView1.Visible = False
End Sub

Private Sub DataGrid1_DblClick()
If IdCliente <> 0 Then
IdCliente = DataGrid1.Columns(0).Value
Check3.Value = 1
End If
Select Case botao

            Case 1
            IdCliente = DataGrid1.Columns(0).Value
            NomeCliente = DataGrid1.Columns(3).Value
            Check3.Value = 1
            Check5.Value = 1
             Combo2.Text = "Informações anexadas!"
            Case 2 ' verificar se ja obteve cliente e se nao podera obter cliente data numero do pedido ,caso contrario receber somente numpedido e data
            If IdCliente = 0 Then
             IdCliente = DataGrid1.Columns(19).Value
             Check3.Value = 1
             IdCliente = DataGrid1.Columns(19).Value
             numPedido = DataGrid1.Columns(8).Value
             Telefonecliente = DataGrid1.Columns(5).Value
             NomeCliente = DataGrid1.Columns(3).Value
             Operadores = DataGrid1.Columns(13).Value
             Combo2.Text = "Nº pedido encontrado!"
      
             Combo1.Text = "Telefone encontrado!"
             
             'BairroCliente As String
             
             Else
                     End If
                    
                    numPedido = DataGrid1.Columns(2).Value
                    DataPedido = DataGrid1.Columns(19).Value
                    Check4.Value = 1
                    Check5.Value = 1
            Case 3 'verificado itens do pedido
            descProdutos = DataGrid1.Columns(2).Value
            Operadores = DataGrid1.Columns(4).Value
                 numPedido = DataGrid1.Columns(6).Value
                 DataPedido = DataGrid1.Columns(5).Value
                NomeCliente = DataGrid1.Columns(7).Value
                Combo1.Text = "Nome Encontrado!"
                 Check3.Value = 1
    End Select
    Check2.Value = 0
    Check1.Value = 0
    consultar = 0
End Sub

Private Sub Form_Load()
preechercombos

End Sub


Public Sub preechercombos()
'combo Cliente
Combo1.Text = "Opções..."
Combo1.AddItem "Telefone"
Combo1.AddItem "Nome"
Combo1.AddItem "Bairro"

'combo Pedidos feitos
Combo2.Text = "Opções..."
Combo2.AddItem "Operadores"
Combo2.AddItem "Produtos"
Combo2.AddItem "Nº Pedido"
Combo2.AddItem "Nome Cliente"
'combo imprimir cancelar
Combo3.Text = "Opções..."
Combo3.AddItem "Reeprimir Cupon"
Combo3.AddItem "Cancelar pedido"
Combo3.AddItem "Relatórios(Versão Adm)"


End Sub

Public Sub IniciarconsultaCliente(indexOptConsultaCliente As Integer)
ConServer


Dim sql As String
'Dim InformacaoDoCliente As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset
Set rs.ActiveConnection = con
rs.CursorLocation = adUseClient

        'listar clientes
        'sql =
        'rs.Open sql
        
        'indices
'0"Telefone"
'1"Nome"
'2 "Bairro"

Select Case indexOptConsultaCliente 'inicia a cosulta distinta ao cliente !
            Case 0 'por Telefone
          
           sql = "SELECT * FROM `Cli_clientes` WHERE `telefone` LIKE '" & InformacaoDoCliente & "' ORDER BY `Cli_clientes`.`id` DESC"
           rs.Open sql
                        If rs.BOF = False Then
                          Set DataGrid1.DataSource = rs
                        FormatDataGrid
                        Else
                        MsgBox " " & consultaLocal & "Não Localizado ! ", , "Não foi possível retornar " & Combo1.Text
                        
                        Combo1.Visible = True
                        
                        End If
                        
                        
                        
                        
                        
            Case 1 'por nome
             sql = "SELECT * FROM `Cli_clientes` WHERE `nome` LIKE '" & InformacaoDoCliente & "' ORDER BY `Cli_clientes`.`id` DESC"
           rs.Open sql
                        If rs.BOF = False Then
                          Set DataGrid1.DataSource = rs
                        FormatDataGrid
                          Else
                        MsgBox " " & consultaLocal & "Não Localizado ! ", , "Não foi possível retornar " & Combo1.Text
                        
                        Combo1.Visible = True
                        
                        End If
                        
         
            Case 2 ' por Bairro
                         sql = "SELECT * FROM `Cli_clientes` WHERE `bairro` LIKE '" & InformacaoDoCliente & "' ORDER BY `Cli_clientes`.`id` DESC"
           rs.Open sql
                        If rs.BOF = False Then
                          Set DataGrid1.DataSource = rs
                        FormatDataGrid
                          Else
                        MsgBox " " & consultaLocal & "Não Localizado ! ", , "Não foi possível retornar " & Combo1.Text
                        
                        Combo1.Visible = True
                        
                        End If
            
            Case 3
            Case 4
   
            Case 5
            Case 5
    End Select
    Check2.Value = 0
    Check1.Value = 0
    consultar = 0







 Set rs = Nothing




End Sub
Public Sub IniciarconsultaPedido(indexOptConsultaCliente As Integer)
ConServer


Dim sql As String
'Dim InformacaoDoCliente As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset
Set rs.ActiveConnection = con
rs.CursorLocation = adUseClient

        'listar clientes
        'sql =
        'rs.Open sql
        
        'indices
'        Combo2.AddItem "Operadores"
'        Combo2.AddItem "Produtos"
'        Combo2.AddItem "Nº Pedido"


Select Case indexOptConsultaCliente 'inicia a cosulta distinta ao cliente !
            Case 0 'por operador
            
           sql = "SELECT * FROM `at_Cupon` WHERE `operador` LIKE '" & InformacaoDoCliente & "'ORDER BY `at_Cupon`.`id` DESC "
           rs.Open sql
                        If rs.BOF = False Then
                          Set DataGrid1.DataSource = rs
                        FormatDataGrid2
                        Else
                        MsgBox " " & consultaLocal & " Não Localizado ! ", , "Não foi possível retornar " & Combo2.Text
                        
                        Combo2.Visible = True
                        
                        End If
                        
                        
                        
                        
                        
            Case 1 'Descrição por itens de um pedido
            
                        sql = "SELECT * FROM `at_itens` WHERE `descrição` LIKE '" & InformacaoDoCliente & "'ORDER BY `id` DESC"
                        rs.Open sql
                        If rs.BOF = False Then
                          Set DataGrid1.DataSource = rs
                          FormatDataGrid3
                          Else
                          MsgBox " " & consultaLocal & " Produto não  Localizado ! ", , "Não foi possível retornar " & Combo2.Text
                        
                          Combo2.Visible = True
                        
                        End If
                        
         
            Case 2 ' numero do pedido
                        sql = "SELECT * FROM `at_itens` WHERE `fk_pedido` = '" & consultaLocalpedidos & "'ORDER BY `id` DESC"
                        rs.Open sql
                        If rs.BOF = False Then
                          Set DataGrid1.DataSource = rs
                          Check6.Value = 1
                         FormatDataGrid3
                          Else
                        MsgBox " " & consultaLocal & " Não Localizado ! ", , "Não foi possível retornar " & Combo2.Text
                        
                        Combo2.Visible = True
                        
                        End If
            
            Case 3 ' nome do cliente
            
                                    sql = "SELECT * FROM `at_itens` WHERE `fk_cliente` LIKE  '" & "%" + consultaLocalpedidos + "%" & "'ORDER BY `id` DESC"
                        rs.Open sql
                        If rs.BOF = False Then
                          Set DataGrid1.DataSource = rs
                          Check6.Value = 1
                         FormatDataGrid3
                          Else
                        MsgBox " " & consultaLocal & " Não Localizado ! ", , "Não foi possível retornar " & Combo2.Text
                        
                        Combo2.Visible = True
                        
                        End If
            
            
            Case 4
   
            Case 5
            Case 5
    End Select
    Check2.Value = 0
    Check1.Value = 0
    consultar = 0







 Set rs = Nothing




End Sub
Public Sub FormatDataGrid()
         
           'cabeçalho do Grid
           DataGrid1.Caption = "Retorno de consulta " & Combo1.Text
           'cabeçalho de colunas"
           DataGrid1.Columns(0).Caption = "Código"
           DataGrid1.Columns(1).Caption = "Telefone"
           DataGrid1.Columns(2).Caption = "CEP"
            DataGrid1.Columns(3).Caption = "Nome"
             DataGrid1.Columns(4).Caption = "Endereço"
              DataGrid1.Columns(5).Caption = "Nº"
               DataGrid1.Columns(6).Caption = "Complemento"
                DataGrid1.Columns(7).Caption = "Apt / Bloco"
                 DataGrid1.Columns(8).Caption = "Referêcia "
                 DataGrid1.Columns(9).Caption = "Bairro"
                 DataGrid1.Columns(10).Caption = "UF"
                 DataGrid1.Columns(11).Caption = "Cidade"
                 DataGrid1.Columns(12).Caption = "Loja Pedido"
End Sub
Public Sub FormatDataGrid2()
         
           'cabeçalho do Grid
           DataGrid1.Caption = "Retorno de consulta"
           'cabeçalho de colunas"
           DataGrid1.Columns(0).Caption = "Código"
           DataGrid1.Columns(1).Caption = "Empresa"
           DataGrid1.Columns(2).Caption = "Nº"
            DataGrid1.Columns(3).Caption = "Nome do Cliente"
             DataGrid1.Columns(4).Caption = "Endereço"
              DataGrid1.Columns(5).Caption = "Telefone"
               DataGrid1.Columns(6).Caption = "Referência"
                DataGrid1.Columns(7).Caption = "Loja"
                 DataGrid1.Columns(8).Caption = "Produtos"
                 DataGrid1.Columns(9).Caption = "Valor Frete"
                 DataGrid1.Columns(10).Caption = "Observações"
                 DataGrid1.Columns(11).Caption = "Total"
                 DataGrid1.Columns(12).Caption = "Data Hora"
                 DataGrid1.Columns(13).Caption = "Operador"
                  DataGrid1.Columns(14).Caption = "Recebido"
                   DataGrid1.Columns(15).Caption = "Valor Pago"
                    DataGrid1.Columns(16).Caption = "Troco"
                     DataGrid1.Columns(17).Caption = "Observações de preparo"
                      DataGrid1.Columns(18).Caption = "Forma de Pagamento"
                       DataGrid1.Columns(19).Caption = "Nº Cliente"
                        
                       
                      
End Sub
Public Sub FormatDataGrid3()
         
           'cabeçalho do Grid
           DataGrid1.Caption = "Retorno de consulta por intens de um pedido"
           'cabeçalho de colunas"
           DataGrid1.Columns(0).Caption = "Código"
           DataGrid1.Columns(1).Caption = "Qtde"
           DataGrid1.Columns(2).Caption = "Descrição"
            DataGrid1.Columns(3).Caption = "valor"
             DataGrid1.Columns(4).Caption = "Operador"
              DataGrid1.Columns(5).Caption = "Data hora"
               DataGrid1.Columns(6).Caption = "Código do pedido"
                DataGrid1.Columns(7).Caption = "Nome do cliente"

                        
                       
                      
End Sub

Private Sub MonthView1_DateDblClick(ByVal DateDblClicked As Date)
Dim resp As Integer

resp = MsgBox("Consultar referências na data  " & DateDblClicked & "?", vbYesNo, "Consultar por data")
If resp = 6 Then
consultaLocal = Trim(DateDblClicked)

   
   dataconsulta = Trim("%" + consultaLocal + "%")
   consultarPorData
   End If


End Sub

Public Sub consultarPorData()
'SELECT * FROM `at_Cupon` WHERE `datahora` LIKE '%03/09/2019%' ORDER BY `fk_Cliente` DESC

ConServer


Dim sql As String
'Dim InformacaoDoCliente As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset
Set rs.ActiveConnection = con
rs.CursorLocation = adUseClient
  sql = "SELECT * FROM `at_Cupon` WHERE `datahora` LIKE '" & dataconsulta & "' ORDER BY `fk_Cliente` DESC"
 ' sql = "SELECT * FROM `at_Cupon` WHERE `datahora` LIKE '%03/09/2019%' ORDER BY `fk_Cliente` DESC"
           rs.Open sql
                        If rs.BOF = False Then
                          Set DataGrid1.DataSource = rs
                        FormatDataGrid2
                        MonthView1.Visible = False
                          Else
                        MsgBox " " & consultaLocal & " Não Localizada ! ", , "Não foi possível retornar valores desta data"
                        
                        MonthView1.Visible = True
                        
                        End If
 Set rs = Nothing
                        

End Sub

Public Sub contadordePassos(passo As Integer)
Select Case passo
'contar em que botao foi clicado para a devida operaçao ser feita
            Case 1 'aberto seleção para cliente
            botao = 1
            Case 2 'varrer as datas
            botao = 2
            Case 10
            'nome entrada
            Case 2
            Case 22
            'data nome entrada e saida
            Case 4
    End Select
    Check2.Value = 0
    Check1.Value = 0
    consultar = 0
End Sub

Private Sub MonthView1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MonthView1.Visible = True
End Sub

Public Function FuncaoRetornevalorDefalt(indexcombo As Integer) As String

Select Case indexcombo
'Combo1.AddItem "Telefone"
'Combo1.AddItem "Nome"
'Combo1.AddItem "Bairro"
            Case 0 'aberto seleção para cliente
               FuncaoRetornevalorDefalt = Telefonecliente
      
 
            Case 1 'varrer as datas
                       FuncaoRetornevalorDefalt = NomeCliente
            Combo1.Text = "Nome Encontrado!"
            Case 2 'nome entrada
            
            
            
    End Select

End Function


Public Function FuncaoRetornevalorDefaltcombo2(indexcombo As Integer) As String

Select Case indexcombo
'Combo2.AddItem "Operadores"
'Combo2.AddItem "Produtos"
'Combo2.AddItem "Nº Pedido"
            Case 0 'aberto seleção para cliente
               FuncaoRetornevalorDefaltcombo2 = Operadores
               Combo2.Text = "Valores encotrados"
      
 
            Case 1 'varrer as datas
                       
               FuncaoRetornevalorDefaltcombo2 = numPedido
            Combo2.Text = "Valores encotrados"
            Case 2 'nome entrada
            FuncaoRetornevalorDefaltcombo2 = numPedido
            Combo2.Text = "Valores encotrados"
            
                        Case 3 'nome entrada
                        Check4.Value = 1
            FuncaoRetornevalorDefaltcombo2 = Trim(NomeCliente)
            Combo2.Text = "Valores encotrados"
            
    End Select

End Function


Public Sub avancadocombo()

''combo Cliente
'Combo1.Text = "Opções..."
'Combo1.AddItem "Telefone"
'Combo1.AddItem "Nome"
'Combo1.AddItem "Bairro"
'
''combo Pedidos feitos
'Combo2.Text = "Opções..."
'Combo2.AddItem "Operadores"
'Combo2.AddItem "Produtos"
'Combo2.AddItem "Nº Pedido"
'Combo2.AddItem "Nome Cliente"
''combo imprimir cancelar
'Combo3.Text = "Opções..."
'Combo3.AddItem "Reeprimir Cupon"
'Combo3.AddItem "Cancelar pedido"
'Combo3.AddItem "Relatórios(Versão Adm)"


End Sub
