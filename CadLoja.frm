VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form13 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configurar impressora"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11070
   Icon            =   "CadLoja.frx":0000
   LinkTopic       =   "Form13"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   11070
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3975
      Index           =   1
      Left            =   480
      TabIndex        =   40
      Top             =   1800
      Width           =   9375
      Begin VB.Timer Timer2 
         Left            =   7560
         Top             =   1920
      End
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   120
         TabIndex        =   47
         Top             =   1560
         Visible         =   0   'False
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   661
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Timer Timer1 
         Left            =   5280
         Top             =   1440
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Command9"
         Height          =   495
         Left            =   8280
         TabIndex        =   46
         Top             =   840
         Width           =   615
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Definir impressora padrão"
         Height          =   375
         Left            =   5400
         TabIndex        =   45
         Top             =   2760
         Width           =   3735
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Testar impressora"
         Height          =   375
         Left            =   840
         TabIndex        =   44
         Top             =   1080
         Width           =   3735
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   8760
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Configurar impressão"
         Height          =   375
         Left            =   720
         TabIndex        =   43
         Top             =   2760
         Width           =   3735
      End
      Begin VB.ComboBox cboTemp 
         Height          =   315
         Left            =   840
         TabIndex        =   42
         Text            =   "Combo1"
         Top             =   720
         Width           =   6735
      End
      Begin VB.Label Label14 
         Caption         =   "Impressora"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   41
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3975
      Index           =   0
      Left            =   480
      TabIndex        =   22
      Top             =   1680
      Width           =   9375
      Begin VB.TextBox Text40 
         DataField       =   "frete"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   5880
         TabIndex        =   7
         Text            =   "Text40"
         Top             =   3480
         Width           =   3255
      End
      Begin VB.TextBox Text1 
         DataField       =   "razao_social"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   360
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   720
         Width           =   4575
      End
      Begin VB.TextBox Text2 
         DataField       =   "nome_fantazia"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   5520
         TabIndex        =   23
         Text            =   "Text2"
         Top             =   720
         Width           =   3615
      End
      Begin VB.TextBox Text3 
         DataField       =   "enderco"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Text            =   "Text3"
         Top             =   1440
         Width           =   4815
      End
      Begin VB.TextBox Text4 
         DataField       =   "numero"
         DataSource      =   "Adodc1"
         Height          =   405
         Left            =   5520
         TabIndex        =   2
         Text            =   "Text4"
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox Text5 
         DataField       =   "complemento"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   6480
         TabIndex        =   3
         Text            =   "Text5"
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox Text6 
         DataField       =   "bairro"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Text            =   "Text6"
         Top             =   2160
         Width           =   2055
      End
      Begin VB.TextBox Text7 
         DataField       =   "cidade"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   2520
         TabIndex        =   5
         Text            =   "Text7"
         Top             =   2160
         Width           =   2055
      End
      Begin VB.TextBox Text9 
         DataField       =   "uf"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   6480
         TabIndex        =   8
         Text            =   "Text9"
         Top             =   2160
         Width           =   495
      End
      Begin VB.TextBox Text10 
         DataField       =   "contato"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   7200
         TabIndex        =   9
         Text            =   "Text10"
         Top             =   2160
         Width           =   1935
      End
      Begin VB.TextBox Text14 
         DataField       =   "site"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   5880
         TabIndex        =   13
         Text            =   "Text14"
         Top             =   2880
         Width           =   1575
      End
      Begin VB.TextBox Text15 
         DataField       =   "email"
         DataSource      =   "Adodc1"
         Height          =   405
         Left            =   7560
         TabIndex        =   14
         Text            =   "Text15"
         Top             =   2880
         Width           =   1575
      End
      Begin VB.TextBox Text8 
         DataField       =   "cep"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   4680
         TabIndex        =   6
         Text            =   "Text8"
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox Text11 
         DataField       =   "telefone"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   360
         TabIndex        =   10
         Text            =   "Text11"
         Top             =   2880
         Width           =   1815
      End
      Begin VB.TextBox Text12 
         DataField       =   "fax"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   2280
         TabIndex        =   11
         Text            =   "Text12"
         Top             =   2880
         Width           =   1695
      End
      Begin VB.TextBox Text13 
         DataField       =   "celular"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   4080
         TabIndex        =   12
         Text            =   "Text13"
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Label Label11 
         Caption         =   "Frete"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   5160
         TabIndex        =   48
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label Label20 
         Caption         =   "Nome Fantazia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   5520
         TabIndex        =   39
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label15 
         Caption         =   "Razão Social"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   38
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label13 
         Caption         =   "Email"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   7560
         TabIndex        =   37
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Label12 
         Caption         =   "Celular"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   4200
         TabIndex        =   36
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "Site"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   5880
         TabIndex        =   35
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "Fax"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   2280
         TabIndex        =   34
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Telefone"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   33
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Contato"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   7200
         TabIndex        =   32
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "CEP "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   4680
         TabIndex        =   31
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "UF"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   6480
         TabIndex        =   30
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Cidade"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   2520
         TabIndex        =   29
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Bairro "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   28
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Complemento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   6480
         TabIndex        =   27
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Nº"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   5520
         TabIndex        =   26
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Endereço"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   25
         Top             =   1200
         Width           =   855
      End
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   4935
      Left            =   120
      TabIndex        =   21
      Top             =   1080
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   8705
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Registro"
            Object.Tag             =   ""
            ImageVarType    =   2
            ImageIndex      =   6
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Configuração"
            Object.Tag             =   ""
            ImageVarType    =   2
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdMoveLast 
      Caption         =   "Mover"
      Enabled         =   0   'False
      Height          =   615
      Left            =   7680
      Picture         =   "CadLoja.frx":0A02
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton CmdMoveFrist 
      Caption         =   "Mover"
      Enabled         =   0   'False
      Height          =   615
      Left            =   1800
      Picture         =   "CadLoja.frx":1404
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton CmdExcluir 
      Caption         =   "Excluir"
      Enabled         =   0   'False
      Height          =   615
      Left            =   6720
      Picture         =   "CadLoja.frx":1E06
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "Editar"
      Height          =   615
      Left            =   4800
      Picture         =   "CadLoja.frx":2808
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton CmdConsultar 
      Caption         =   "Consultar"
      CausesValidation=   0   'False
      Height          =   615
      Left            =   3720
      Picture         =   "CadLoja.frx":320A
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdNovo 
      Caption         =   "Novo"
      Height          =   615
      Left            =   2640
      Picture         =   "CadLoja.frx":3C0C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdSalvar 
      Caption         =   "Salvar"
      Enabled         =   0   'False
      Height          =   615
      Left            =   5760
      Picture         =   "CadLoja.frx":460E
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   120
      Width           =   855
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   11880
      Top             =   1560
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=robofi61_order_taker"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "robofi61_order_taker"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "cadastro_da_empresa"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   9000
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   15
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CadLoja.frx":5010
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CadLoja.frx":51EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CadLoja.frx":53C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CadLoja.frx":559E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CadLoja.frx":5778
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CadLoja.frx":5952
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CadLoja.frx":5B2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CadLoja.frx":5D06
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CadLoja.frx":5EE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CadLoja.frx":60BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CadLoja.frx":6294
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CadLoja.frx":646E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CadLoja.frx":6648
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CadLoja.frx":6822
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CadLoja.frx":69FC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   1680
      X2              =   8400
      Y1              =   840
      Y2              =   840
   End
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Impressoras As Printer
Dim Pvalor, valor As Integer
Dim arraYtextos(25) As String



Private Sub Command1_Click()
'prncnfg: Configure ou mostra as configurações de impressora.
'prndrvr: Adiciona, deleta e lista drivers de impressoras.
'prnjobs: Pausa, continua ou cancela lista de documentos a serem impressos.
'prnmngr: Adiciona, deleta e lista impressoras conectadas, além da impressora default.
'prnport: Cria, deleta e lista portas de impressora TCP/IP
'prnqctl: Imprime uma página de teste, pausa ou reinicia um documento a ser impresso
'PrintQueue(PrintServer, String)


End Sub

Private Sub CmdConsultar_Click()
Dim nomeUserConsultar As String
nomeUserConsultar = InputBox("Entre com bairro da loja ", "Consultar")
ConsultarUsuario (nomeUserConsultar)
End Sub

Private Sub cmdEditar_Click()
cmdSalvar.Enabled = True

End Sub

Private Sub CmdExcluir_Click()
Adodc1.Recordset.Delete
End Sub

Private Sub Command2_Click()
CommonDialog1.ShowPrinter
End Sub

Private Sub Command3_Click()
Dim TESTE As String
Dim BOLI As Boolean

'TESTE = Impressoras.Circle

'BOLI = Impressoras.Copies
'TESTE = Impressoras.EndDoc
BOLI = Impressoras.hDC
'TESTE = Impressoras.KillDoc
'TESTE = Impressoras.line
'TESTE = Impressoras.NewPage
'TESTE = Impressoras.PaintPicture
'TESTE = Impressoras.PSet
TESTE = Impressoras.TrackDefault
TESTE = Impressoras.PrintQuality






End Sub

Private Sub Command4_Click()
cmdImprimir
Command4.Enabled = False
Timer1.Interval = 50000
Timer2.Interval = 5000
ProgressBar1.Value = 0
valor = 0
ProgressBar1.Visible = True


End Sub

Private Sub Command5_Click()
  Dim PaginaInicial, Paginafinal, numerodecopias, i
    CommonDialog1.CancelError = True
    On Error GoTo TrataErro

    'mostra a janela para impressora
    CommonDialog1.ShowPrinter
    'Captura os valores definidos pelo usuário na janela
    PaginaInicial = CommonDialog1.FromPage
    Paginafinal = CommonDialog1.ToPage
    numerodecopias = CommonDialog1.Copies
    For i = 1 To numerodecopias
        'aqui entra o seu código para imprimir
    Next
    Exit Sub

TrataErro:
    Exit Sub

End Sub

Private Sub Command6_Click()
Dim command As String
'command = "c:\windows\notepad.exe"
command = "CONTROL PRINTERS"
Shell "cmd.exe /c " & command
End Sub

Private Sub Command8_Click()
CommonDialog1.ShowColor

End Sub

Private Sub Form_Load()
'C:\Windows\System32\spool\PRINTERS
 ajusta_container
 PreencheImpressora


 fecheformAS


End Sub
Public Sub ConsultarUsuario(nome As String)
If nome <> "" Then


Adodc1.RecordSource = ""
nome = Trim(nome + "%")
Adodc1.CommandType = adCmdText
Adodc1.RecordSource = "SELECT * FROM `cadastro_da_empresa` WHERE `bairro`  LIKE '" & nome & "'"

Adodc1.Refresh
 


If (Adodc1.Recordset.EOF = False) Then



   
       
'    TabStrip1.Tabs(2).Selected = True
'    Frame1(1).Visible = True
'    Frame1(0).Visible = False
'    Else
'    TabStrip1.Tabs(1).Selected = True
'     valiDacao = 0
'
'    Frame1(0).Visible = True
'    Frame1(1).Visible = False
    Else
  nome = Replace(nome, "%", "")
   MsgBox "Não foi possivel encontra a loja  : " & nome & " ", , "Loja não encontrada"
  
End If
End If
End Sub



Private Sub TabStrip1_Click()
Dim i As Integer

i = TabStrip1.SelectedItem.Index

Frame1(i - 1).ZOrder

  

End Sub
Private Sub cmdImprimir()

For Each Impressoras In Printers
    Set Printer = Impressoras
    If Impressoras.DeviceName = cboTemp.Text Then Exit For
Next

Printer.Print "Robofild!" & _
"Tele bh Amarelinho!" & _
"Teste!"
Printer.EndDoc

End Sub
Private Sub ajusta_container()
Dim i As Integer
With TabStrip1
For i = 1 To .Tabs.Count
Frame1(i - 1).Move .ClientLeft, .ClientTop, .ClientWidth, .ClientHeight
Next
End With

TabStrip1.Tabs(1).Selected = True
   TabStrip1_Click
       
    
End Sub
Public Sub PreencheImpressora()
Dim prtImpressora As Printer

  On Error GoTo PreencheImpressora_Error

cboTemp.Clear

For Each prtImpressora In Printers
  cboTemp.AddItem prtImpressora.DeviceName
Next

Set prtImpressora = Nothing

  On Error GoTo 0
  Exit Sub

PreencheImpressora_Error:

  'MsgBox  Error   & Err.Number &   (  & Err.Description &  ) in procedure PreencheImpressora of Módulo basImprimeDoctos
  Err.Clear
  
End Sub

Public Sub SetImpressoraLocalSistema(ByVal IMPRESSORA As String)
Dim X      As Printer
Dim Y      As String
    Y = UCase$(IMPRESSORA)
    For Each X In Printers
        If UCase$(X.DeviceName) = Y Then
            Set Printer = X
            Exit For
        End If
    Next X
End Sub

Private Sub Text40_Change()
If Text40.Text <> "" Then

'Text40.Text = Format(Text40.Text, "Currency")
End If
End Sub

Private Sub Text40_LostFocus()
If Text40.Text <> "" Then
Dim testo40 As Double
    Text40.Text = Text40.Text
 
testo40 = CDbl(Text40.Text)
Text40.Text = Format(Text40.Text, "Currency")
End If
End Sub

Private Sub Timer1_Timer()

 
If VerifiqueImpressão = True Then
Timer1.Interval = 0
Timer2.Interval = 0
Command4.Enabled = True
Else
Timer1.Interval = 0
Timer2.Interval = 0
Command4.Enabled = True
End If

End Sub

Private Sub Timer2_Timer()
ProgressBar1.Visible = True

Pvalor = valor + Pvalor
valor = 10
ProgressBar1.Value = Pvalor
If ProgressBar1.Value = 100 Then
ProgressBar1.Visible = False
Timer2.Interval = 0
End If

End Sub
'integracao
'Option Explicit

Private Sub cmdNovo_Click()
capturando
cmdSalvar.Enabled = True

Adodc1.Recordset.AddNew
abraCadabra
devolver
Text3.SetFocus
'DataCombo1.Text = "Defina o Regime"
End Sub

Private Sub cmdSalvar_Click()
Adodc1.Recordset.Update
Adodc1.Recordset.MoveFirst
Adodc1.Refresh
MsgBox "Salvo com sucesso!", , "Salvo"
End Sub

Private Sub DataCombo1_Change()
'If DataCombo1.Text <> "Defina o Regime" Then
'DataCombo1.Enabled = False
'
'End If
End Sub



Private Sub MaskEdBox1_LostFocus()
'Text32.Text = MaskEdBox1.Text
End Sub

Private Sub MaskEdBox2_LostFocus()
'Dim QuardeValorphone As String
'QuardeValorphone = MaskEdBox2.Text

'Select Case constroimaskparaTelefone(MaskEdBox2.Text)
 '           Case 11
  '           MaskEdBox2.Mask = "(##)#####-####"
   '           MaskEdBox2.Text = QuardeValorphone
              
    '        Case 9
      '       MsgBox "completo so com 9"
     '        'constroimaskparaTelefone = 9
       '     Case 8
        '    'constroimaskparaTelefone = 8
          '    MsgBox "completo sem o 9"
         '   Case Else
            'constroimaskparaTelefone = 0
           '    MsgBox "simples "
    'End Select






End Sub




Public Function constroimaskparaTelefone(fone As String)



End Function

Private Sub Text1_LostFocus()
Text1.Text = Format(Text1.Text, ">")

End Sub

Private Sub Text16_GotFocus()


'Text16.SetFocus
End Sub

Private Sub Text16_KeyPress(KeyAscii As Integer)
KeyAscii = (SoNumeros(KeyAscii))
   If KeyAscii = 0 Then
   End If
End Sub

'Private Sub Text17_LostFocus()
'Text17.Text = Format(Text17.Text, ">")
'End Sub
'
'Private Sub Text18_LostFocus()
'Text18.Text = Format(Text18.Text, ">")
'End Sub
'
'Private Sub Text19_LostFocus()
'Text19.Text = Format(Text19.Text, ">")
'End Sub

Private Sub Text2_LostFocus()
Text2.Text = Format(Text2.Text, ">")

End Sub

'Private Sub Text21_LostFocus()
'Text21.Text = Format(Text21.Text, ">")
'End Sub
'Private Sub Text23_LostFocus()
'Text23.Text = Format(Text23.Text, "<")
'End Sub
'
'Private Sub Text22_LostFocus()
'Dim numero As String
'
'numero = RemoverCaracter(Text22.Text)
' If Len(numero) > 0 Then
'      Select Case Len(numero)
'       Case Is = 11
'
'         If Not calculacpf(numero) Then
'            MsgBox "CPF incorreto !!!"
'            'foco
'            Text22.SetFocus
'
'
'            Else
'            Text22.Text = CpfFormat(numero)
'
'         End If
'       Case Is = 14
'
'         If Not ValidaCGC(numero) Then
'            MsgBox "CGC incorreto !!! "
'            'foco
'              Text22.SetFocus
'             Else
'            Text22.Text = CnpjFormat(numero)
'
'         End If
'      End Select
'    End If
'End Sub
'
'Private Sub Text24_LostFocus()
'Text24.Text = phoneformat(Text24.Text)
'
'TabStrip1.Tabs(3).Selected = True
'   TabStrip1_Click
'End Sub

'Private Sub Text28_LostFocus()
'Text28.Text = Format(Text28.Text, ">")
'cmdSalvar.Enabled = True
'cmdSalvar.SetFocus
'
'End Sub

Private Sub Text3_LostFocus()
Text3.Text = Format(Text3.Text, ">")
'Text25.Text = Text3.Text


End Sub

Private Sub Text4_LostFocus()
Text4.Text = Format(Text4.Text, ">")
'Text26.Text = Text4.Text


End Sub
Private Sub Text5_LostFocus()
Text5.Text = Format(Text5.Text, ">")
    'Text27.Text = Text5.Text

End Sub
Private Sub Text6_LostFocus()
Text6.Text = Format(Text6.Text, ">")
'Text30.Text = Text6.Text


End Sub
Private Sub Text7_LostFocus()
Text7.Text = Format(Text7.Text, ">")
'Text31.Text = Text7.Text


End Sub
Private Sub Text9_LostFocus()
Text9.Text = Format(Text9.Text, ">")

'Text33.Text = Text9.Text
End Sub
Private Sub Text10_LostFocus()
Text10.Text = Format(Text10.Text, ">")
'Text34.Text = Text10.Text

End Sub
Private Sub Text14_LostFocus()
Text14.Text = Format(Text14.Text, "<")
'Text38.Text = Text14.Text

End Sub
Private Sub Text15_LostFocus()
Text15.Text = Format(Text15.Text, "<")
'Text39.Text = Text15.Text
 'TabStrip1.Tabs(2).Selected = True
   


End Sub
'Private Sub Text16_LostFocus()
'Dim numero As String
'
'numero = RemoverCaracter(Text16.Text)
' If Len(numero) > 0 Then
'      Select Case Len(numero)
'       Case Is = 11
'
'         If Not calculacpf(numero) Then
'            MsgBox "CPF incorreto !!!"
'            'foco
'            Text16.SetFocus
'
'
'            Else
'            Text16.Text = CpfFormat(numero)
'
'         End If
'       Case Is = 14
'
'         If Not ValidaCGC(numero) Then
'            MsgBox "CGC incorreto !!! "
'            'foco
'              Text16.SetFocus
'             Else
'            Text16.Text = CnpjFormat(numero)
'
'         End If
'      End Select
'    End If
'    Text29.Text = Text16.Text
'
'
'End Sub

Private Sub Text11_LostFocus()
Text11.Text = phoneformat(Text11.Text)
'Text35.Text = Text11.Text

End Sub

Private Sub Text12_LostFocus()
Text12.Text = phoneformat(Text12.Text)
'Text36.Text = Text12.Text
End Sub

Private Sub Text13_Lostfocus()
Text13.Text = phoneformat(Text13.Text)
'Text37.Text = Text13.Text
End Sub

Private Sub Text8_LostFocus()
Text8.Text = CepMask(Text8.Text)

   ' Text32.Text = Text8.Text

End Sub


Public Sub fecheformAS()



If Text1.Text <> "" Then
Text1.Enabled = False
End If

If Text2.Text <> "" Then
Text2.Enabled = False
End If

If Text3.Text <> "" Then
Text3.Enabled = False
End If

If Text4.Text <> "" Then
Text4.Enabled = False
End If

If Text5.Text <> "" Then
Text5.Enabled = False
End If

If Text6.Text <> "" Then
Text6.Enabled = False
End If

If Text7.Text <> "" Then
Text7.Enabled = False
End If

If Text8.Text <> "" Then
Text8.Enabled = False
End If

If Text9.Text <> "" Then
Text9.Enabled = False
End If

If Text10.Text <> "" Then
Text10.Enabled = False
End If

If Text11.Text <> "" Then
Text11.Enabled = False
End If


If Text12.Text <> "" Then
Text12.Enabled = False
End If

If Text13.Text <> "" Then
Text13.Enabled = False
End If

If Text14.Text <> "" Then
Text14.Enabled = False
End If

If Text15.Text <> "" Then
Text15.Enabled = False
End If




'If Text17.Text <> "" Then
'Text17.Enabled = False
'End If
'
'
'If Text18.Text <> "" Then
'Text18.Enabled = False
'End If
'
'If Text19.Text <> "" Then
'Text19.Enabled = False
'End If
'
'
'If Text21.Text <> "" Then
'Text21.Enabled = False
'End If
'
'
'If Text22.Text <> "" Then
'Text22.Enabled = False
'End If
'
'
'If Text23.Text <> "" Then
'Text23.Enabled = False
'End If
'
'
'If Text24.Text <> "" Then
'Text24.Enabled = False
'End If


End Sub


Public Sub abraCadabra()




Text1.Enabled = True


Text2.Enabled = True


Text3.Enabled = True



Text4.Enabled = True



Text5.Enabled = True


Text6.Enabled = True



Text7.Enabled = True


Text8.Enabled = True


Text9.Enabled = True



Text10.Enabled = True



Text11.Enabled = True




Text12.Enabled = True



Text13.Enabled = True



Text14.Enabled = True



Text15.Enabled = True





'If Text17.Text <> "" Then
'Text17.Enabled = False
'End If
'
'
'If Text18.Text <> "" Then
'Text18.Enabled = False
'End If
'
'If Text19.Text <> "" Then
'Text19.Enabled = False
'End If
'
'
'If Text21.Text <> "" Then
'Text21.Enabled = False
'End If
'
'
'If Text22.Text <> "" Then
'Text22.Enabled = False
'End If
'
'
'If Text23.Text <> "" Then
'Text23.Enabled = False
'End If
'
'
'If Text24.Text <> "" Then
'Text24.Enabled = False
'End If


End Sub

Public Sub capturando()
arraYtextos(1) = Text1.Text
arraYtextos(2) = Text2.Text
arraYtextos(9) = Text9.Text
arraYtextos(10) = Text10.Text
arraYtextos(11) = Text11.Text
arraYtextos(12) = Text12.Text
arraYtextos(13) = Text13.Text
arraYtextos(14) = Text14.Text
arraYtextos(15) = Text15.Text


End Sub

Public Sub devolver()
 Text1.Text = arraYtextos(1)
 Text2.Text = arraYtextos(2)
 Text9.Text = arraYtextos(9)
  Text10.Text = arraYtextos(10)
 Text11.Text = arraYtextos(11)
 Text12.Text = arraYtextos(12)
 Text13.Text = arraYtextos(13)
 Text14.Text = arraYtextos(14)
 Text15.Text = arraYtextos(15)


End Sub


