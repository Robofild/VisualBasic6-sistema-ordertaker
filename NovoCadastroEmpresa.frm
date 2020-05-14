VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form12 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro da Empresa"
   ClientHeight    =   6600
   ClientLeft      =   5835
   ClientTop       =   3300
   ClientWidth     =   10200
   Icon            =   "NovoCadastroEmpresa.frx":0000
   LinkTopic       =   "Form12"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   10200
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdMoveLast 
      Caption         =   "Mover"
      Enabled         =   0   'False
      Height          =   615
      Left            =   6240
      Picture         =   "NovoCadastroEmpresa.frx":0A02
      Style           =   1  'Graphical
      TabIndex        =   74
      Top             =   360
      Width           =   735
   End
   Begin VB.CommandButton CmdMoveFrist 
      Caption         =   "Mover"
      Enabled         =   0   'False
      Height          =   615
      Left            =   360
      Picture         =   "NovoCadastroEmpresa.frx":1404
      Style           =   1  'Graphical
      TabIndex        =   73
      Top             =   360
      Width           =   735
   End
   Begin VB.CommandButton CmdExcluir 
      Caption         =   "Excluir"
      Enabled         =   0   'False
      Height          =   615
      Left            =   5280
      Picture         =   "NovoCadastroEmpresa.frx":1E06
      Style           =   1  'Graphical
      TabIndex        =   72
      Top             =   360
      Width           =   855
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "Editar"
      Enabled         =   0   'False
      Height          =   615
      Left            =   3360
      Picture         =   "NovoCadastroEmpresa.frx":2808
      Style           =   1  'Graphical
      TabIndex        =   71
      Top             =   360
      Width           =   855
   End
   Begin VB.CommandButton CmdConsultar 
      Caption         =   "Consultar"
      CausesValidation=   0   'False
      Enabled         =   0   'False
      Height          =   615
      Left            =   2280
      Picture         =   "NovoCadastroEmpresa.frx":320A
      Style           =   1  'Graphical
      TabIndex        =   70
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton cmdNovo 
      Caption         =   "Novo"
      Enabled         =   0   'False
      Height          =   615
      Left            =   1200
      Picture         =   "NovoCadastroEmpresa.frx":3C0C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton cmdSalvar 
      Caption         =   "Salvar"
      Height          =   615
      Left            =   4320
      Picture         =   "NovoCadastroEmpresa.frx":460E
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   360
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Height          =   3375
      Index           =   0
      Left            =   480
      TabIndex        =   54
      Top             =   2280
      Width           =   9375
      Begin VB.TextBox Text13 
         DataField       =   "celular"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   4080
         TabIndex        =   13
         Text            =   "Text13"
         Top             =   2880
         Width           =   1695
      End
      Begin VB.TextBox Text12 
         DataField       =   "fax"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   2280
         TabIndex        =   12
         Text            =   "Text12"
         Top             =   2880
         Width           =   1695
      End
      Begin VB.TextBox Text11 
         DataField       =   "telefone"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   360
         TabIndex        =   11
         Text            =   "Text11"
         Top             =   2880
         Width           =   1815
      End
      Begin VB.TextBox Text8 
         DataField       =   "cep"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   4680
         TabIndex        =   8
         Text            =   "Text8"
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox Text15 
         DataField       =   "email"
         DataSource      =   "Adodc1"
         Height          =   405
         Left            =   7560
         TabIndex        =   15
         Text            =   "Text15"
         Top             =   2880
         Width           =   1575
      End
      Begin VB.TextBox Text14 
         DataField       =   "site"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   5880
         TabIndex        =   14
         Text            =   "Text14"
         Top             =   2880
         Width           =   1575
      End
      Begin VB.TextBox Text10 
         DataField       =   "contato"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   7200
         TabIndex        =   10
         Text            =   "Text10"
         Top             =   2160
         Width           =   1935
      End
      Begin VB.TextBox Text9 
         DataField       =   "uf"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   6480
         TabIndex        =   9
         Text            =   "Text9"
         Top             =   2160
         Width           =   495
      End
      Begin VB.TextBox Text7 
         DataField       =   "cidade"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   2520
         TabIndex        =   7
         Text            =   "Text7"
         Top             =   2160
         Width           =   2055
      End
      Begin VB.TextBox Text6 
         DataField       =   "bairro"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Text            =   "Text6"
         Top             =   2160
         Width           =   2055
      End
      Begin VB.TextBox Text5 
         DataField       =   "complemento"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   6480
         TabIndex        =   5
         Text            =   "Text5"
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox Text4 
         DataField       =   "numero"
         DataSource      =   "Adodc1"
         Height          =   405
         Left            =   5520
         TabIndex        =   4
         Text            =   "Text4"
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox Text3 
         DataField       =   "enderco"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Text            =   "Text3"
         Top             =   1440
         Width           =   4815
      End
      Begin VB.TextBox Text2 
         DataField       =   "nome_fantazia"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   5520
         TabIndex        =   2
         Text            =   "Text2"
         Top             =   720
         Width           =   3615
      End
      Begin VB.TextBox Text1 
         DataField       =   "razao_social"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   720
         Width           =   4575
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
         TabIndex        =   69
         Top             =   1200
         Width           =   855
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
         TabIndex        =   68
         Top             =   1200
         Width           =   495
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
         TabIndex        =   67
         Top             =   1200
         Width           =   1455
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
         TabIndex        =   66
         Top             =   1920
         Width           =   855
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
         TabIndex        =   65
         Top             =   1920
         Width           =   735
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
         TabIndex        =   64
         Top             =   1920
         Width           =   735
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
         TabIndex        =   63
         Top             =   1920
         Width           =   1095
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
         TabIndex        =   62
         Top             =   1920
         Width           =   1455
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
         TabIndex        =   61
         Top             =   2640
         Width           =   855
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
         TabIndex        =   60
         Top             =   2640
         Width           =   735
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
         TabIndex        =   59
         Top             =   2640
         Width           =   735
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
         TabIndex        =   58
         Top             =   2640
         Width           =   1095
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
         TabIndex        =   57
         Top             =   2640
         Width           =   1455
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
         TabIndex        =   56
         Top             =   480
         Width           =   1335
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
         TabIndex        =   55
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3375
      Index           =   1
      Left            =   480
      TabIndex        =   43
      Top             =   2400
      Width           =   8775
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   330
         Left            =   360
         Top             =   3000
         Visible         =   0   'False
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   582
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
         RecordSource    =   "Cad_Tributario"
         Caption         =   "Adodc2"
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
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "NovoCadastroEmpresa.frx":5010
         DataField       =   "regime_tributario"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   360
         TabIndex        =   20
         Top             =   2520
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "nome_tributario"
         Text            =   "DataCombo1"
      End
      Begin VB.TextBox Text19 
         DataField       =   "cnae"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   360
         TabIndex        =   19
         Text            =   "Text19"
         Top             =   1440
         Width           =   2295
      End
      Begin VB.TextBox Text18 
         DataField       =   "insc_municipal_identidade"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   5880
         TabIndex        =   18
         Text            =   "Text18"
         Top             =   720
         Width           =   2535
      End
      Begin VB.TextBox Text17 
         DataField       =   "insc_estadual_identidade"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   3000
         TabIndex        =   17
         Text            =   "Text17"
         Top             =   720
         Width           =   2535
      End
      Begin VB.TextBox Text16 
         DataField       =   "cnpj_cpf"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   360
         TabIndex        =   16
         Text            =   "Text16"
         Top             =   720
         Width           =   2295
      End
      Begin VB.Frame Frame3 
         Caption         =   "Dados do Contador "
         Height          =   1935
         Left            =   3120
         TabIndex        =   44
         Top             =   1320
         Width           =   5655
         Begin VB.TextBox Text24 
            DataField       =   "telefone_contador"
            DataSource      =   "Adodc1"
            Height          =   375
            Left            =   2760
            TabIndex        =   24
            Text            =   "Text24"
            Top             =   1320
            Width           =   2175
         End
         Begin VB.TextBox Text23 
            DataField       =   "Email_contador"
            DataSource      =   "Adodc1"
            Height          =   375
            Left            =   240
            TabIndex        =   23
            Text            =   "Text23"
            Top             =   1320
            Width           =   2415
         End
         Begin VB.TextBox Text22 
            DataField       =   "cpf_cnpj_contador"
            DataSource      =   "Adodc1"
            Height          =   375
            Left            =   2760
            TabIndex        =   22
            Text            =   "Text22"
            Top             =   600
            Width           =   2175
         End
         Begin VB.TextBox Text21 
            DataField       =   "nome_contador"
            DataSource      =   "Adodc1"
            Height          =   375
            Left            =   240
            TabIndex        =   21
            Text            =   "Text21"
            Top             =   600
            Width           =   2415
         End
         Begin VB.Label Label14 
            Caption         =   "Nome"
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
            TabIndex        =   48
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label TXTcPFcONT 
            AutoSize        =   -1  'True
            Caption         =   "CPF/CNPJ"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2760
            TabIndex        =   47
            Top             =   360
            Width           =   915
         End
         Begin VB.Label Label19 
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
            Left            =   240
            TabIndex        =   46
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label Label18 
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
            Left            =   2880
            TabIndex        =   45
            Top             =   1080
            Width           =   855
         End
      End
      Begin VB.Label Label22 
         Caption         =   "Regime Tributário"
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
         Left            =   360
         TabIndex        =   53
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Label23 
         Caption         =   "CNAE"
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
         Left            =   360
         TabIndex        =   52
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label24 
         Caption         =   "INSC.Municipal"
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
         Left            =   5880
         TabIndex        =   51
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "INSC.ESTADUAL/IDENTIDADE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2880
         TabIndex        =   50
         Top             =   480
         Width           =   2715
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "CPF/CNPJ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   49
         Top             =   480
         Width           =   915
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4095
      Index           =   2
      Left            =   360
      TabIndex        =   27
      Top             =   1800
      Width           =   9495
      Begin VB.TextBox Text39 
         DataField       =   "recibo_email"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   7560
         TabIndex        =   89
         Text            =   "Text39"
         Top             =   3000
         Width           =   1575
      End
      Begin VB.TextBox Text38 
         DataField       =   "recibo_site"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   6000
         TabIndex        =   88
         Text            =   "Text38"
         Top             =   3000
         Width           =   1455
      End
      Begin VB.TextBox Text37 
         DataField       =   "recibo_celular"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   4320
         TabIndex        =   87
         Text            =   "Text37"
         Top             =   3000
         Width           =   1575
      End
      Begin VB.TextBox Text36 
         DataField       =   "recibo_fax"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   2400
         TabIndex        =   86
         Text            =   "Text36"
         Top             =   3000
         Width           =   1815
      End
      Begin VB.TextBox Text35 
         DataField       =   "recibo_telefone"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   360
         TabIndex        =   85
         Text            =   "Text35"
         Top             =   3000
         Width           =   1815
      End
      Begin VB.TextBox Text34 
         DataField       =   "recibo_contato"
         DataSource      =   "Adodc1"
         Height          =   405
         Left            =   7320
         TabIndex        =   84
         Text            =   "Text34"
         Top             =   2160
         Width           =   1815
      End
      Begin VB.TextBox Text33 
         DataField       =   "recibo_uf"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   6480
         TabIndex        =   83
         Text            =   "Text33"
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox Text32 
         DataField       =   "recibo_cep"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   4680
         TabIndex        =   82
         Text            =   "Text32"
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox Text31 
         DataField       =   "Recibo_cidade"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   2640
         TabIndex        =   81
         Text            =   "Text31"
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox Text30 
         DataField       =   "recibo_bairro"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   360
         TabIndex        =   80
         Text            =   "Text30"
         Top             =   2160
         Width           =   1935
      End
      Begin VB.TextBox Text29 
         DataField       =   "recibo_cnpj"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   5520
         TabIndex        =   79
         Text            =   "Text29"
         Top             =   1200
         Width           =   3495
      End
      Begin VB.TextBox Text28 
         DataField       =   "recibo_referencia"
         DataSource      =   "Adodc1"
         Height          =   405
         Left            =   360
         TabIndex        =   25
         Text            =   "Text28"
         Top             =   1200
         Width           =   4815
      End
      Begin VB.TextBox Text27 
         DataField       =   "recibo_complemento"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   6480
         TabIndex        =   78
         Text            =   "Text27"
         Top             =   480
         Width           =   2415
      End
      Begin VB.TextBox Text26 
         DataField       =   "recibo_numero"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   5400
         TabIndex        =   77
         Text            =   "Text26"
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox Text25 
         DataField       =   "recibo_nome"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   360
         TabIndex        =   76
         Text            =   "Text25"
         Top             =   480
         Width           =   4815
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
         Index           =   1
         Left            =   7440
         TabIndex        =   42
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
         Index           =   1
         Left            =   4200
         TabIndex        =   41
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
         Index           =   1
         Left            =   5880
         TabIndex        =   40
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
         Index           =   1
         Left            =   2280
         TabIndex        =   39
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
         Index           =   1
         Left            =   360
         TabIndex        =   38
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
         Index           =   1
         Left            =   7200
         TabIndex        =   37
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
         Index           =   1
         Left            =   4680
         TabIndex        =   36
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
         Index           =   1
         Left            =   6480
         TabIndex        =   35
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
         Index           =   1
         Left            =   2520
         TabIndex        =   34
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
         Index           =   1
         Left            =   360
         TabIndex        =   33
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
         Index           =   1
         Left            =   6480
         TabIndex        =   32
         Top             =   240
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
         Index           =   1
         Left            =   5520
         TabIndex        =   31
         Top             =   240
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
         Index           =   1
         Left            =   360
         TabIndex        =   30
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label16 
         Caption         =   "Referência"
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
         Left            =   360
         TabIndex        =   29
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "CPF/CNPJ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5520
         TabIndex        =   28
         Top             =   960
         Width           =   915
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   7080
      Top             =   840
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
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   4335
      Left            =   240
      TabIndex        =   75
      Top             =   1560
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   7646
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Dados Cadastrais"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
            ImageIndex      =   6
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Dados Fiscais"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Texto Para Recibos"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
            ImageIndex      =   8
         EndProperty
      EndProperty
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   6720
      Y1              =   1320
      Y2              =   1320
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   7440
      Top             =   240
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
            Picture         =   "NovoCadastroEmpresa.frx":5025
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "NovoCadastroEmpresa.frx":51FF
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "NovoCadastroEmpresa.frx":53D9
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "NovoCadastroEmpresa.frx":55B3
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "NovoCadastroEmpresa.frx":578D
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "NovoCadastroEmpresa.frx":5967
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "NovoCadastroEmpresa.frx":5B41
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "NovoCadastroEmpresa.frx":5D1B
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "NovoCadastroEmpresa.frx":5EF5
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "NovoCadastroEmpresa.frx":60CF
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "NovoCadastroEmpresa.frx":62A9
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "NovoCadastroEmpresa.frx":6483
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "NovoCadastroEmpresa.frx":665D
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "NovoCadastroEmpresa.frx":6837
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "NovoCadastroEmpresa.frx":6A11
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdNovo_Click()
Adodc1.Recordset.AddNew
DataCombo1.Text = "Defina o Regime"
End Sub

Private Sub cmdSalvar_Click()
Adodc1.Recordset.Update
Adodc1.Recordset.MoveFirst
Adodc1.Refresh
End Sub

Private Sub DataCombo1_Change()
If DataCombo1.Text <> "Defina o Regime" Then
DataCombo1.Enabled = False

End If
End Sub

Private Sub Form_Load()
 ajusta_container
 fecheformAS
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

Private Sub TabStrip1_Click()
Dim i As Integer

i = TabStrip1.SelectedItem.Index

Frame1(i - 1).ZOrder

  

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

Private Sub Text17_LostFocus()
Text17.Text = Format(Text17.Text, ">")
End Sub

Private Sub Text18_LostFocus()
Text18.Text = Format(Text18.Text, ">")
End Sub

Private Sub Text19_LostFocus()
Text19.Text = Format(Text19.Text, ">")
End Sub

Private Sub Text2_LostFocus()
Text2.Text = Format(Text2.Text, ">")

End Sub

Private Sub Text21_LostFocus()
Text21.Text = Format(Text21.Text, ">")
End Sub
Private Sub Text23_LostFocus()
Text23.Text = Format(Text23.Text, "<")
End Sub

Private Sub Text22_LostFocus()
Dim numero As String

numero = RemoverCaracter(Text22.Text)
 If Len(numero) > 0 Then
      Select Case Len(numero)
       Case Is = 11
        
         If Not calculacpf(numero) Then
            MsgBox "CPF incorreto !!!"
            'foco
            Text22.SetFocus
           
            
            Else
            Text22.Text = CpfFormat(numero)
           
         End If
       Case Is = 14
         
         If Not ValidaCGC(numero) Then
            MsgBox "CGC incorreto !!! "
            'foco
              Text22.SetFocus
             Else
            Text22.Text = CnpjFormat(numero)
           
         End If
      End Select
    End If
End Sub

Private Sub Text24_LostFocus()
Text24.Text = phoneformat(Text24.Text)

TabStrip1.Tabs(3).Selected = True
   TabStrip1_Click
End Sub

Private Sub Text28_LostFocus()
Text28.Text = Format(Text28.Text, ">")
cmdSalvar.Enabled = True
cmdSalvar.SetFocus

End Sub

Private Sub Text3_LostFocus()
Text3.Text = Format(Text3.Text, ">")
Text25.Text = Text3.Text


End Sub

Private Sub Text4_LostFocus()
Text4.Text = Format(Text4.Text, ">")
Text26.Text = Text4.Text


End Sub
Private Sub Text5_LostFocus()
Text5.Text = Format(Text5.Text, ">")
Text27.Text = Text5.Text

End Sub
Private Sub Text6_LostFocus()
Text6.Text = Format(Text6.Text, ">")
Text30.Text = Text6.Text


End Sub
Private Sub Text7_LostFocus()
Text7.Text = Format(Text7.Text, ">")
Text31.Text = Text7.Text


End Sub
Private Sub Text9_LostFocus()
Text9.Text = Format(Text9.Text, ">")

Text33.Text = Text9.Text
End Sub
Private Sub Text10_LostFocus()
Text10.Text = Format(Text10.Text, ">")
Text34.Text = Text10.Text

End Sub
Private Sub Text14_LostFocus()
Text14.Text = Format(Text14.Text, "<")
Text38.Text = Text14.Text

End Sub
Private Sub Text15_LostFocus()
Text15.Text = Format(Text15.Text, "<")
Text39.Text = Text15.Text
 TabStrip1.Tabs(2).Selected = True
   


End Sub
Private Sub Text16_LostFocus()
Dim numero As String

numero = RemoverCaracter(Text16.Text)
 If Len(numero) > 0 Then
      Select Case Len(numero)
       Case Is = 11
        
         If Not calculacpf(numero) Then
            MsgBox "CPF incorreto !!!"
            'foco
            Text16.SetFocus
           
            
            Else
            Text16.Text = CpfFormat(numero)
           
         End If
       Case Is = 14
         
         If Not ValidaCGC(numero) Then
            MsgBox "CGC incorreto !!! "
            'foco
              Text16.SetFocus
             Else
            Text16.Text = CnpjFormat(numero)
           
         End If
      End Select
    End If
    Text29.Text = Text16.Text
    

End Sub

Private Sub Text11_LostFocus()
Text11.Text = phoneformat(Text11.Text)
Text35.Text = Text11.Text

End Sub

Private Sub Text12_LostFocus()
Text12.Text = phoneformat(Text12.Text)
Text36.Text = Text12.Text
End Sub

Private Sub Text13_Lostfocus()
Text13.Text = phoneformat(Text13.Text)
Text37.Text = Text13.Text
End Sub

Private Sub Text8_LostFocus()
Text8.Text = CepMask(Text8.Text)

    Text32.Text = Text8.Text

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


If Text16.Text <> "" Then
Text16.Enabled = False
End If


If Text17.Text <> "" Then
Text17.Enabled = False
End If


If Text18.Text <> "" Then
Text18.Enabled = False
End If

If Text19.Text <> "" Then
Text19.Enabled = False
End If


If Text21.Text <> "" Then
Text21.Enabled = False
End If


If Text22.Text <> "" Then
Text22.Enabled = False
End If


If Text23.Text <> "" Then
Text23.Enabled = False
End If


If Text24.Text <> "" Then
Text24.Enabled = False
End If


End Sub
