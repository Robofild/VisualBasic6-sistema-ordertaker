VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.ocx"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "comct232.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Atendimento "
   ClientHeight    =   11085
   ClientLeft      =   975
   ClientTop       =   1155
   ClientWidth     =   20880
   Icon            =   "Atendimento.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   11085
   ScaleWidth      =   20880
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      Caption         =   "Bairro"
      Height          =   855
      Left            =   5640
      TabIndex        =   73
      Top             =   3480
      Visible         =   0   'False
      Width           =   4215
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Clique mais uma vez para confirmar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   600
         TabIndex        =   75
         Top             =   120
         Width           =   3015
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Clique para ver o Bairro"
         Height          =   195
         Left            =   240
         TabIndex        =   74
         Top             =   360
         Width           =   1650
      End
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   9000
      TabIndex        =   71
      Text            =   "Text7"
      Top             =   240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command13 
      Enabled         =   0   'False
      Height          =   375
      Left            =   5640
      Picture         =   "Atendimento.frx":0A02
      Style           =   1  'Graphical
      TabIndex        =   70
      Top             =   3000
      Width           =   375
   End
   Begin MSDataListLib.DataList DataList1 
      Height          =   5820
      Left            =   720
      TabIndex        =   2
      Top             =   3360
      Visible         =   0   'False
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   10266
      _Version        =   393216
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame3 
      Caption         =   "Wathsapp"
      Height          =   9015
      Index           =   1
      Left            =   11520
      TabIndex        =   68
      Top             =   2400
      Width           =   8415
      Begin SHDocVwCtl.WebBrowser WebBrowser2 
         Height          =   8655
         Left            =   120
         TabIndex        =   69
         Top             =   240
         Width           =   8175
         ExtentX         =   14420
         ExtentY         =   15266
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   1
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   375
      Left            =   7800
      Picture         =   "Atendimento.frx":1404
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   64
      Top             =   600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   7680
      Picture         =   "Atendimento.frx":1E06
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   63
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Command10"
      Height          =   375
      Left            =   8520
      TabIndex        =   61
      Top             =   3000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      Caption         =   "loc frete"
      Height          =   375
      Left            =   13080
      TabIndex        =   59
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FF8080&
      Caption         =   "Alterar Frete"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7320
      MousePointer    =   3  'I-Beam
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   4440
      Width           =   1095
   End
   Begin VB.TextBox Text16 
      Height          =   375
      Left            =   10320
      TabIndex        =   52
      Text            =   "Text16"
      Top             =   4680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text15 
      Height          =   375
      Left            =   9720
      TabIndex        =   51
      Text            =   "C:\Order_Taker\FINDFILE.AVI"
      Top             =   0
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox Text14 
      Height          =   375
      Left            =   19680
      TabIndex        =   49
      Text            =   "C:\Order_Taker\FINDFILE.AVI"
      Top             =   15480
      Width           =   2655
   End
   Begin VB.TextBox Text13 
      Height          =   285
      Left            =   11160
      TabIndex        =   48
      Text            =   "Text13"
      Top             =   360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   375
      Left            =   9480
      TabIndex        =   47
      Top             =   5280
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "F1               Pedido"
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
      Height          =   615
      Left            =   720
      TabIndex        =   11
      Top             =   6120
      Width           =   6855
   End
   Begin VB.Timer Timer2 
      Left            =   9840
      Top             =   3720
   End
   Begin VB.Timer Timer1 
      Left            =   9840
      Top             =   3120
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Apartamento?"
      Height          =   375
      Left            =   9360
      TabIndex        =   46
      Top             =   2400
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalvar 
      Caption         =   "Salvar"
      Enabled         =   0   'False
      Height          =   615
      Left            =   4800
      Picture         =   "Atendimento.frx":2248
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton cmdNovo 
      Caption         =   "Novo"
      Enabled         =   0   'False
      Height          =   615
      Left            =   1680
      Picture         =   "Atendimento.frx":2C4A
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton CmdConsultar 
      Caption         =   "Consultar"
      CausesValidation=   0   'False
      Enabled         =   0   'False
      Height          =   615
      Left            =   2760
      Picture         =   "Atendimento.frx":364C
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "Editar"
      Enabled         =   0   'False
      Height          =   615
      Left            =   3840
      Picture         =   "Atendimento.frx":404E
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton CmdLimpar 
      Caption         =   "Limpar"
      Enabled         =   0   'False
      Height          =   615
      Left            =   5760
      Picture         =   "Atendimento.frx":4A50
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton CmdLojas 
      Caption         =   "Lojas"
      Height          =   615
      Left            =   720
      Picture         =   "Atendimento.frx":5452
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox Text12 
      Height          =   495
      Left            =   9120
      TabIndex        =   38
      Text            =   "Text12"
      Top             =   1440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text11 
      Height          =   495
      Left            =   9120
      TabIndex        =   37
      Text            =   "inicio"
      Top             =   720
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   9240
      Top             =   4320
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      RecordSource    =   "cadastro_da_empresa"
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
   Begin VB.TextBox Text10 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Left            =   1800
      TabIndex        =   8
      Text            =   "BELO HORIZONTE"
      Top             =   4440
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Left            =   720
      TabIndex        =   7
      Text            =   "MG"
      Top             =   4440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   3720
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   3000
      Visible         =   0   'False
      Width           =   4935
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   8400
      Top             =   1920
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
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Order_Taker\CEP.MDB;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Order_Taker\CEP.MDB;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "CEP"
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
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   6000
      TabIndex        =   6
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   3720
      Width           =   2895
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Atendimento.frx":5E54
      Height          =   360
      Left            =   4320
      TabIndex        =   9
      Top             =   4440
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   635
      _Version        =   393216
      ListField       =   "bairro"
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "Atendimento.frx":5E69
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   2160
      Width           =   6495
   End
   Begin VB.TextBox txtCep 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3120
      TabIndex        =   12
      Top             =   1320
      Width           =   1815
   End
   Begin VB.TextBox txtRua 
      DataField       =   "ENDERECO"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   33
      Top             =   3000
      Width           =   4935
   End
   Begin VB.TextBox txtnumero 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6360
      TabIndex        =   3
      Top             =   3000
      Width           =   855
   End
   Begin VB.TextBox txtbairro 
      DataField       =   "BAIRRO"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   32
      Top             =   3720
      Width           =   1935
   End
   Begin VB.TextBox txtEstado 
      DataField       =   "UF"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   31
      Top             =   4440
      Width           =   615
   End
   Begin VB.TextBox txtCidade 
      DataField       =   "CIDADE"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1800
      TabIndex        =   30
      Top             =   4440
      Width           =   2175
   End
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      Height          =   9015
      Index           =   0
      Left            =   11400
      TabIndex        =   28
      Top             =   1080
      Width           =   8415
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         Height          =   9015
         Left            =   -120
         TabIndex        =   29
         Top             =   0
         Width           =   8535
         ExtentX         =   15055
         ExtentY         =   15901
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   1
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   9975
      Left            =   11040
      TabIndex        =   27
      Top             =   360
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   17595
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Google Maps"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Use os fantásticos recursos google para sua auditoria"
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Whatsapp"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Caption         =   "Referencia para entrega "
      Height          =   975
      Left            =   720
      TabIndex        =   26
      Top             =   5040
      Width           =   6855
      Begin VB.TextBox Text2 
         Height          =   495
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   240
         Width           =   6495
      End
   End
   Begin VB.CommandButton Command4 
      Height          =   375
      Left            =   5040
      Picture         =   "Atendimento.frx":687B
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Enabled         =   0   'False
      Height          =   375
      Left            =   7320
      Picture         =   "Atendimento.frx":727D
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "**Cancelar**"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8160
      Picture         =   "Atendimento.frx":7C7F
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   8880
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pedidos"
      Height          =   3375
      Left            =   360
      TabIndex        =   22
      Top             =   6840
      Visible         =   0   'False
      Width           =   10455
      Begin VB.CommandButton Command8 
         Caption         =   "F4   Reabrir pedido "
         Height          =   855
         Left            =   7800
         Picture         =   "Atendimento.frx":8681
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   1200
         Width           =   2295
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Voltar aos últimos pedidos"
         Height          =   855
         Left            =   7800
         Picture         =   "Atendimento.frx":9083
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   360
         Visible         =   0   'False
         Width           =   2295
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2535
         Left            =   360
         TabIndex        =   56
         Top             =   360
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   4471
         _Version        =   393216
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
         Caption         =   "Últimos pedidos deste cliente"
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
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   19
      Text            =   " "
      Top             =   1320
      Width           =   2055
   End
   Begin ComCtl2.Animation Animation1 
      Height          =   855
      Left            =   7680
      TabIndex        =   50
      Top             =   5880
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1508
      _Version        =   327681
      FullWidth       =   81
      FullHeight      =   57
   End
   Begin VB.CommandButton cmdEntrega 
      Caption         =   "Entrega"
      Enabled         =   0   'False
      Height          =   615
      Left            =   6840
      Picture         =   "Atendimento.frx":9A85
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label18 
      Caption         =   "Label18"
      Height          =   255
      Left            =   9000
      TabIndex        =   72
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label16 
      Caption         =   "&F5"
      Height          =   375
      Left            =   360
      TabIndex        =   66
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Label15 
      Caption         =   "&F3"
      Height          =   375
      Left            =   360
      TabIndex        =   65
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "Valor de frete recomendado =>"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   4320
      TabIndex        =   62
      Top             =   4800
      Visible         =   0   'False
      Width           =   2190
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "Label13"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   6600
      TabIndex        =   60
      Top             =   4800
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Label Label12 
      Caption         =   "Complemento"
      Height          =   255
      Left            =   2880
      TabIndex        =   55
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label11 
      Caption         =   "retorn indice"
      Height          =   375
      Left            =   9960
      TabIndex        =   54
      Top             =   720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label10 
      Caption         =   "bloco"
      Height          =   255
      Left            =   6480
      TabIndex        =   36
      Top             =   3480
      Width           =   615
   End
   Begin VB.Label Label9 
      Caption         =   "Apt\"
      Height          =   255
      Left            =   6120
      TabIndex        =   35
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label Label8 
      Caption         =   "Selecione loja para entrega"
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
      Left            =   4440
      TabIndex        =   34
      Top             =   4200
      Width           =   2775
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   19320
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   3
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Atendimento.frx":A487
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Atendimento.frx":A661
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Atendimento.frx":A83B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label7 
      Caption         =   "Cidade "
      Height          =   255
      Left            =   1800
      TabIndex        =   25
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Label U 
      Caption         =   "UF"
      Height          =   255
      Left            =   720
      TabIndex        =   24
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label Label6 
      Caption         =   "Bairro"
      Height          =   255
      Left            =   720
      TabIndex        =   21
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Nº"
      Height          =   255
      Left            =   6240
      TabIndex        =   20
      Top             =   2760
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Telefone"
      Height          =   255
      Left            =   720
      TabIndex        =   18
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Endereço:"
      Height          =   255
      Left            =   720
      TabIndex        =   17
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "CEP"
      Height          =   255
      Left            =   3120
      TabIndex        =   16
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Nome"
      Height          =   255
      Left            =   720
      TabIndex        =   15
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label17 
      Caption         =   "&F2"
      Height          =   255
      Left            =   6960
      TabIndex        =   67
      Top             =   0
      Width           =   255
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim telefoneTipo As Integer
Dim valorRuaAnterior As String
Dim chaveMestra As Integer
Dim chaveEdicao As Integer
Dim numPedido As Integer
Dim clintesemalteraçoes As Boolean
Dim gridselecionado As Boolean
Dim bdEntreganull As String
Dim TROCAimagem As Boolean
Dim buscainicialAiEntrega As Boolean
Dim consultaDinamica As Boolean
Dim COSULTAR As Boolean
Dim varText8 As String
Dim varText6 As String
Dim EnvieManual As Boolean


  


 


Private Sub Combo1_Click()


End Sub

Private Sub CmdConsultar_Click()
Text11.Text = "inicio"
Text12.Text = "inicio"
ConsultarTelefone
'Command2.Enabled = True

End Sub

Private Sub cmdEditar_Click()
clintesemalteraçoes = True
'Command2.Enabled = True
resgateid
chaveMestra = 0
AbiliteBloco (chaveMestra)
chaveEdicao = 5
End Sub

Private Sub cmdEntrega_Click()
'animação

If TROCAimagem = False Then
buscainicialAiEntrega = True
Form401.Text33.Text = 1
Form404.Text11.Text = 1
Aibuscar
Else
'entregar
buscainicialAiEntrega = False
Form401.Text33.Text = 2
Form404.Text11.Text = 2
AiEntrega
End If


'Form406_1.Show

End Sub

Private Sub CmdLimpar_Click()
LimparAtendimento
End Sub

Private Sub cmdMoveLast_Click()

End Sub

Private Sub CmdLojas_Click()
Form25.Show
Form25.DataCombo2.Text = Text8.Text
End Sub

Private Sub cmdNovo_Click()
'clintesemalteraçoes = True
'CmdLimpar.Enabled = True
''Command2.Enabled = True
'chaveMestra = 0
'chaveEdicao = 0
'AbiliteBloco (chaveMestra)
'Text1.Text = ""
'Text1.SetFocus

clintesemalteraçoes = True
CmdLimpar.Enabled = True
'Command2.Enabled = True
chaveMestra = 0
chaveEdicao = 0
Text11 = "inicio"
AbiliteBloco (chaveMestra)
Text1.Text = ""
Text1.SetFocus
End Sub

Private Sub cmdSalvar_Click()
If DataCombo1.Text <> "" And DataCombo1.Text <> "Selecione a Loja!" Then
REMOVespacosEmBranco
CaixaAlta

ConServer

Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset
Set rs.ActiveConnection = con

If chaveEdicao <> 1 And Text11 = "inicio" Then
'entrada de dados
If tXTcEP = "" Then
   
    sql = "INSERT INTO `robofi61_order_taker`.`Cli_clientes` (`telefone`, `cep`, `nome`, `endereco`, `numero`, `Apt`, `bloco`, `referecia`, `bairro`, `uf`, `cidade`, `loja`) " & _
    "VALUES (' " & Text4.Text & " ', ' " & tXTcEP.Text & " ',' " & Text1.Text & " ', ' " & Text6.Text & " ', ' " & txtnumero.Text & " ', ' " & Text3.Text & " ',' " & Text5.Text & " ', ' " & Text2.Text & " ', ' " & Text8.Text & " ', ' " & Text9.Text & " ', ' " & Text10.Text & " ', ' " & DataCombo1.Text & " ')"
    rs.Open sql
    Else
    If txtRua = "" Then
    txtRua.Text = Text6.Text
     tXTbAIRRO.Text = Text8.Text
      txtEstado.Text = Text9.Text
       TXTcIDADE.Text = Text10.Text
       End If
    
      sql = "INSERT INTO `robofi61_order_taker`.`Cli_clientes` (`telefone`, `cep`, `nome`, `endereco`, `numero`, `Apt`, `bloco`, `referecia`, `bairro`, `uf`, `cidade`, `loja`) " & _
    "VALUES (' " & Text4.Text & " ', ' " & tXTcEP.Text & " ',' " & Text1.Text & " ', ' " & txtRua.Text & " ', ' " & txtnumero.Text & " ', ' " & Text3.Text & " ',' " & Text5.Text & " ', ' " & Text2.Text & " ', ' " & tXTbAIRRO.Text & " ', ' " & txtEstado.Text & " ', ' " & TXTcIDADE.Text & " ', ' " & DataCombo1.Text & " ')"
    rs.Open sql
    
End If
 Set rs = Nothing
'MsgBox "salvo com sucesso", , "Salvo"
resgateid
Else
'updade no prontuario do cliente
If Text11.Text = "inicio" Then
resgateid
End If
If chaveEdicao = 5 Then

sql = "UPDATE `Cli_clientes` SET `telefone`=' " & Text4.Text & " ',`cep`=' " & tXTcEP.Text & " ',`nome`=' " & Text1.Text & " ',`endereco`=' " & Text6.Text & " ',`numero`= ' " & txtnumero.Text & " ',`Apt`=' " & Text3.Text & " ',`bloco`=' " & Text5.Text & " ',`referecia`=' " & Text2.Text & " ',`bairro`=' " & Text8.Text & " ',`uf`= ' " & Text9.Text & " ',`cidade`=' " & Text10.Text & " ',`loja`=' " & DataCombo1.Text & " ' WHERE `id`=' " & Text11.Text & " '"
Else
sql = "UPDATE `Cli_clientes` SET `telefone`=' " & Text4.Text & " ',`cep`=' " & tXTcEP.Text & " ',`nome`=' " & Text1.Text & " ',`endereco`=' " & Text6.Text & " ',`numero`= ' " & txtnumero.Text & " ',`Apt`=' " & Text3.Text & " ',`bloco`=' " & Text5.Text & " ',`referecia`=' " & Text2.Text & " ',`bairro`=' " & Text8.Text & " ',`uf`= ' " & Text9.Text & " ',`cidade`=' " & Text10.Text & " ' WHERE `id`=' " & Text11.Text & " '"
End If

If chaveEdicao = 5 Then
'MsgBox "Alteração salva com sucesso !", , "Alterado com sucesso"
End If
  rs.Open sql
  Set rs = Nothing
End If
chaveEdicao = 1
End If
End Sub

Private Sub Command1_Click()
Dim resp As Integer
resp = MsgBox("Isso excluirá o pedido por completo deseja continuar?", vbYesNo, "Excluir +++Cancelar")
If resp = 6 Then

Form502.Show
Form502.Label3.Caption = numPedido
CancelamentoPedido
End If
End Sub



Private Sub Command10_Click()
gerarlojaAutomatic
End Sub



Private Sub Command11_Click()

End Sub

Private Sub Command12_Click()
'SELECT `Apt`,`bloco` FROM `Cli_clientes` WHERE `cep` LIKE '30880-420' AND `numero` LIKE '293'
If tXTcEP <> "" And txtnumero <> "" Then
ConServer

Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs.ActiveConnection = con

'localizar  localizar apartamento


   sql = "SELECT `Apt`,`bloco` FROM `Cli_clientes` WHERE `cep` LIKE '" & tXTcEP.Text & "' AND `numero` LIKE '" & txtnumero.Text & "' "
   rs.Open sql
    If rs.BOF = False Then
            If (rs.Fields("Apt").Value <> "") Then
            Timer1.Interval = 500
            End If
             If (rs.Fields("bloco").Value <> "") Then
              Timer2.Interval = 500
            End If
      End If
   rs.Close
   
  
    

 Set rs = Nothing

End If


End Sub

Private Sub Command13_Click()
If DataList1.Visible = False Then

DataList1.Visible = True
consultaDinamica = True
localizarendereco (Text6.Text)
Else
Frame4.Visible = False
DataList1.Visible = False
consultaDinamica = False
End If
'Text6.SetFocus

End Sub



Private Sub Command2_Click()
Form23.Visible = False
Form401.Show
Form401.Visible = False
If cmdEntrega.Enabled = True Then
'CommonDialog1.CancelError = True
'trarar erro
On Error GoTo error

ReverifiqueEntregaouBuscar
Call Command9_Click
'form500.show
''Sleep 2000
Call cmdSalvar_Click
Form500.ProgressBar1.Value = 50

pesquisarFreteassinado

Form500.ProgressBar1.Value = 75
'Form401.Show
Form401.Visible = True
Form401.SetFocus

'revogada
If EnvieManual = True Then
'cliente
Form404.Label12.Caption = Trim(Form2.Text1.Text)
Form404.Label15.Caption = Trim(Form2.Text4.Text)
Form404.Label14.Caption = Trim(Form2.Text2.Text)
Form404.Label13.Caption = Trim(Text6.Text) + ", " + Trim(txtnumero.Text) + "/" + Trim(Text3.Text) + Trim(Text5.Text) + "- " + Trim(Text8.Text) + ", " + Trim(Text10.Text)
Else
Form404.Label12.Caption = Trim(Form2.Text1.Text)
Form404.Label15.Caption = Form2.Text4.Text
Form404.Label14.Caption = Form2.Text2.Text
Form404.Label13.Caption = Trim(txtRua.Text) + ", " + Trim(txtnumero.Text) + "/" + Trim(Text3.Text) + Trim(Text5.Text) + "- " + Trim(tXTbAIRRO.Text) + ", " + Trim(TXTcIDADE.Text)

End If


Form404.Label15.Caption = Form2.Text4.Text
Form404.Label14.Caption = Form2.Text2.Text
Form404.Label17.Caption = Form2.DataCombo1.Text


Else
MsgBox "Selecione a opção de entrega use a TECLA {F2} ", , "Opção de entrega ?"
cmdEntrega.Enabled = True
   Call cmdEntrega_Click
End If

Exit Sub

error:
If cmdEntrega.Enabled = True Then

ReverifiqueEntregaouBuscar
Call Command9_Click
'form500.show
''Sleep 2000
Call cmdSalvar_Click
Form500.ProgressBar1.Value = 50

pesquisarFreteassinado

Form500.ProgressBar1.Value = 75
Form401.Show
Form401.Visible = True
Form401.SetFocus
Else
MsgBox "Selecione a opção de entrega use a TECLA {F2} ", , "Opção de entrega ?"
cmdEntrega.Enabled = True
   Call cmdEntrega_Click
End If


Exit Sub
If cmdEntrega.Enabled = True Then
ReverifiqueEntregaouBuscar
Call Command9_Click
'form500.show
''Sleep 2000
Call cmdSalvar_Click
Form500.ProgressBar1.Value = 50

pesquisarFreteassinado

Form500.ProgressBar1.Value = 75
Form401.Visible = True
Form401.SetFocus
Else
MsgBox "Selecione a opção de entrega use a TECLA {F2} ", , "Opção de entrega ?"
End If

'revogada
If EnvieManual = True Then
'cliente
Form404.Label12.Caption = Trim(Form2.Text1.Text)
Form404.Label15.Caption = Trim(Form2.Text4.Text)
Form404.Label14.Caption = Trim(Form2.Text2.Text)
Form404.Label13.Caption = Trim(Text6.Text) + ", " + Trim(txtnumero.Text) + "/" + Trim(Text3.Text) + Trim(Text5.Text) + "- " + Trim(Text8.Text) + ", " + Trim(Text10.Text)
Else
Form404.Label12.Caption = Trim(Form2.Text1.Text)
Form404.Label15.Caption = Form2.Text4.Text
Form404.Label14.Caption = Form2.Text2.Text
Form404.Label13.Caption = Trim(txtRua.Text) + ", " + Trim(txtnumero.Text) + "/" + Trim(Text3.Text) + Trim(Text5.Text) + "- " + Trim(tXTbAIRRO.Text) + ", " + Trim(TXTcIDADE.Text)

End If

Form404.Label15.Caption = Form2.Text4.Text
Form404.Label14.Caption = Form2.Text2.Text
Form404.Label17.Caption = Form2.DataCombo1.Text

Form404.Label15.Caption = Form2.Text4.Text
Form404.Label14.Caption = Form2.Text2.Text
Form404.Label17.Caption = Form2.DataCombo1.Text

    End Sub

Private Sub Command2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF1 Then
   Call Command2_Click
  End If


End Sub

Private Sub Command3_Click()
On Error GoTo trata_erro
 'retiraCaracteresEspeciais (txtRua.Text)
'retiraCaracteresEspeciais (Text6.Text)

If txtRua <> "" Then
 WebBrowser1.Silent = True
   WebBrowser1.Navigate Trim("http://maps.google.com/maps?q= Rua' " & retiraCaracteresEspeciais(txtRua.Text) & " ',Bairro ' " & retiraCaracteresEspeciais(tXTbAIRRO.Text) & " ',Numero ' " & retiraCaracteresEspeciais(txtnumero.Text) & " ' ,Cidade ' " & retiraCaracteresEspeciais(TXTcIDADE.Text) & " '  ")
Else
 WebBrowser1.Silent = True
   WebBrowser1.Navigate Trim("http://maps.google.com/maps?q= Rua' " & retiraCaracteresEspeciais(Text6.Text) & " ',Bairro ' " & retiraCaracteresEspeciais(Text8.Text) & " ',Numero ' " & retiraCaracteresEspeciais(txtnumero.Text) & " ' ,Cidade ' " & retiraCaracteresEspeciais(Text10.Text) & " '  ")
   End If
   Exit Sub
trata_erro:
 WebBrowser1.Silent = False
    WebBrowser1.Navigate "www.sistemasgratis.com"
  MsgBox Err.Description
'consultaEndereco.Append ("http://maps.google.com/maps?q=")

End Sub

Private Sub Command4_Click()
localizarCep

'https://www.google.com/maps/@-19.904281,-44.022235,16z
On Error GoTo trata_erro
  WebBrowser1.Navigate Trim("https://www.google.com/maps/@-19.904281,-44.022235,16z")
   Exit Sub
trata_erro:
   MsgBox Err.Description
'consultaEndereco.Append ("http://maps.google.com/maps?q=")
End Sub









Private Sub Command8_Click()
'revogada
If EnvieManual = True Then
'cliente
Form404.Label12.Caption = Trim(Form2.Text1.Text)
Form404.Label15.Caption = Trim(Form2.Text4.Text)
Form404.Label14.Caption = Trim(Form2.Text2.Text)
Form404.Label13.Caption = Trim(Text6.Text) + ", " + Trim(txtnumero.Text) + "/" + Trim(Text3.Text) + Trim(Text5.Text) + "- " + Trim(Text8.Text) + ", " + Trim(Text10.Text)
Else
Form404.Label12.Caption = Trim(Form2.Text1.Text)
Form404.Label15.Caption = Form2.Text4.Text
Form404.Label14.Caption = Form2.Text2.Text
Form404.Label13.Caption = Trim(txtRua.Text) + ", " + Trim(txtnumero.Text) + "/" + Trim(Text3.Text) + Trim(Text5.Text) + "- " + Trim(tXTbAIRRO.Text) + ", " + Trim(TXTcIDADE.Text)

End If


Form404.Label15.Caption = Form2.Text4.Text
Form404.Label14.Caption = Form2.Text2.Text
Form404.Label17.Caption = Form2.DataCombo1.Text

Form404.Label15.Caption = Form2.Text4.Text
Form404.Label14.Caption = Form2.Text2.Text
Form404.Label17.Caption = Form2.DataCombo1.Text

If Label13.Caption <> "Label13" Then
Form406.Label4.Caption = Replace(Label13.Caption, ",", ".")
End If
If Text7.Text <> "True" Then

gireoDATACOMBOlojadopedidoeincluaOPERADOR


End If
If gridselecionado = True Then
        Dim resp As Integer
        resp = MsgBox("Deseja mandar um AVISO IMPRESSO  para loja " & DataCombo1.Text & "  informando que vai alterar o pedido !", vbInformation + vbYesNo, "Alteração de pedido ?")
        'form500.show
        If resp = 6 Then
        AlteracaoPedido
        gireoDATACOMBOlojadopedidoeincluaOPERADOR
        End If
        
        verificarsepossuiaFrete
        Call Command9_Click
        reabrirParaedicao (numPedido)
Else
MsgBox "Selecione antes o pedido que deseja reabrir", , "Pedido não selecionado!"
  DataGrid1.SetFocus
  End If
  Form500.Hide
  Form404.Show
End Sub
Public Sub reabrirParaedicao(numPedido As Integer)
Form3.Hide
Dialog.Hide
'form500.show
    ConServer
Dim valorFRete As Double
Dim numerodopedido As Integer
Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs.ActiveConnection = con

Dim valorDosItens As Double
numerodopedido = numPedido
Form500.ProgressBar1.Value = 30
Form404.Show
Form404.Text10 = 1
Form404.Text2.Text = " |-"

Form404.Label20.Caption = "operador"
'data hora
Form404.Label19 = Now
'operador
'Form404.Label20 = ""
'observaçoes
If Text13.Text <> "" Then
Form404.Text2.Text = " |-"
Form500.ProgressBar1.Value = 40
End If
Form404.Caption = "Pedido Nº" & numerodopedido
Form404.Label18.Caption = numerodopedido

Form404.Adodc1.RecordSource = ""

Form404.Adodc1.CommandType = adCmdText

Form404.Adodc1.RecordSource = "SELECT `Quantidade`,`descrição`,`valor` ,`id`FROM `at_itens` WHERE `fk_pedido` = '" & numerodopedido & "'"

Form404.Adodc1.Refresh
'DataGrid1.Columns(0).Caption = NpTitulo
Form404.DataGrid1.Columns(1).Width = 3800
'buscar o frete
Form500.ProgressBar1.Value = 50
Form404.Adodc2.RecordSource = ""

Form404.Adodc2.CommandType = adCmdText

Form404.Adodc2.RecordSource = "SELECT `frete` FROM `at_frete` WHERE `fk_numPedido` ='" & numerodopedido & "'"

Form404.Adodc2.Refresh
Form500.ProgressBar1.Value = 60
If Form404.Adodc2.Recordset.BOF = False And Form404.Label18.Caption <> "0" Then



valorFRete = Format(Form404.Adodc2.Recordset.Fields("frete").Value, "General Number")
Else
MsgBox "Você não encerrou seu pedido tecle (F1) para fechar", , "Fechar o pedido antes de continuar !"
Form404.Hide
Form401.Show

Form500.ProgressBar1.Value = 85
Exit Sub
End If

    If (Form404.Adodc1.Recordset.EOF = False) Then
    


rs.CursorLocation = adUseClient
  sql = "SELECT SUM(`valor`) FROM `at_itens` WHERE `fk_pedido` = '" & numerodopedido & "'"
  rs.Open sql
  valorDosItens = Format(rs.Fields("SUM(`valor`)").Value, "General Number")
  Form404.Text1.Text = Format(valorDosItens + valorFRete, "currency")
    
    Form500.ProgressBar1.Value = 95
    
    
    End If
 Set rs = Nothing
Form500.ProgressBar1.Value = 100
Form404.Show ''

            Set rs = Nothing
 
End Sub
Private Sub Command9_Click()
Dim loja As String
Dim Bairro As String
Bairro = Trim(Text8.Text)
'Bairro = tXTbAIRRO.Text

loja = Trim(DataCombo1.Text)
ConServer


Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs.ActiveConnection = con


rs.CursorLocation = adUseClient
  sql = "SELECT `valor` FROM `re_fretebairroloja` WHERE `bairro` LIKE '" & Bairro & "' AND `loja` LIKE  '%" & loja & "%'  LIMIT 1"
  rs.Open sql
  If rs.BOF = False Then
  Label13.Caption = rs.Fields("valor").Value
  End If
 Set rs = Nothing


End Sub



Private Sub Command6_Click()

Form24.Show


If Text8.Text <> "" Then
Form24.DataCombo2.Text = Trim(Text8.Text)
Else
Form24.DataCombo2.Text = Trim(tXTbAIRRO)
End If
Form24.DataCombo1.Text = DataCombo1.Text
'Command2.Enabled = True
End Sub

Private Sub Command7_Click()
consultarUtimosPedidos (Text11.Text)
Command7.Visible = False

End Sub

Private Sub DataCombo1_Change()
If DataCombo1.Text <> "" And DataCombo1.Text <> " " Then
 Command2.Enabled = True
End If
End Sub

Private Sub DataCombo1_DblClick(Area As Integer)
AiLiberarLojas
End Sub

Private Sub DataGrid1_Click()
gridselecionado = True
Command1.Visible = True
If DataGrid1.Caption <> "Visulização de itens do pedido" Then
numPedido = (DataGrid1.Columns(1).Value)
tragaAlojaDopedido
End If
End Sub

Private Sub DataGrid1_DblClick()
If DataGrid1.Caption <> "Visulização de itens do pedido" Then
DataGrid1.Caption = "Visulização de itens do pedido"
Command8.Enabled = False
consultaritensUtimosPedidos (DataGrid1.Columns(1).Value)
Command7.Visible = True

Command1.Enabled = False
Else
Call Command7_Click
DataGrid1.Caption = "Últimos pedidos deste cliente"
Command7.Visible = False
Command8.Enabled = True
Command1.Enabled = True
End If

End Sub

Private Sub DataList1_DblClick()
Dinamic (DataList1.Text)
'Text6.Text = DataList1.Text

End Sub
Public Function localizarenderecoplist(rua As String)



rua = Trim("%" + rua + "%")
Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs.ActiveConnection = conl
rs.CursorLocation = adUseClient

           sql = "SELECT * FROM CEP WHERE `ENDERECO` LIKE '" & rua & "' "
           rs.Open sql
            If rs.BOF = False Then
            Set Label19.DataSource = rs
             
            Label19.DataField = "BAIRRO"
            'DataCombo1.Text = rs.Fields("loja").Value
           End If
           DataList1.Text = " "

    

 




Set rs = Nothing
End Function

Private Sub DataList1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
localizarenderecoplist (DataList1.Text)

End If
End Sub

Private Sub DataList1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Frame4.Top = Y + 3280
Frame4.Visible = True

End Sub
Private Sub DataList1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Dinamic (DataList1.Text)
End If
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

  If KeyCode = vbKeyF1 Then
   Call Command2_Click
  End If
  If KeyCode = vbKeyF2 Then
  cmdEntrega.Enabled = True
   Call cmdEntrega_Click
  End If

  If KeyCode = vbKeyF3 Then
    
    AtLIberarManual
    Text6.Enabled = True
    If Text6.Visible = True Then
    Text6.SetFocus
    End If
  End If
    If KeyCode = vbKeyF4 Then
    
    Call Command8_Click
    
  End If
  
 
  
     If KeyCode = vbKeyF5 Then
     Text8.Enabled = True
    AtLIberarManual
     Text8.SetFocus
    
  End If
'
End Sub
Private Sub ajusta_container()
Dim i As Integer
With TabStrip1
For i = 1 To .Tabs.Count
Frame3(i - 1).Move .ClientLeft, .ClientTop, .ClientWidth, .ClientHeight
Next
End With

TabStrip1.Tabs(1).Selected = True
   
       
    
End Sub

Private Sub Form_Load()
Form401.Show
Form401.Visible = False
 ajusta_container
Adodc1.RecordSource = ""

Adodc1.CommandType = adCmdText

Adodc1.RecordSource = "SELECT * FROM CEP WHERE CEP ='0' "

Adodc1.Refresh
If Adodc1.Recordset.BOF = False Then

End If
chaveMestra = 1
End Sub

Private Sub Label13_Change()
If Label13.Caption <> "Label13" Then
Form406.Label4.Caption = Replace(Label13.Caption, ",", ".")
End If
If Label13.Caption <> "Label13" Then
Label13.Visible = True
Form406.Label4 = Replace(Label13.Caption, ",", ".")
Else
Label13.Visible = False
End If

End Sub

Private Sub Label19_Change()
If Label20.Visible = False Then
Label20.Visible = True
Else
Label20.Visible = False
End If
End Sub

Private Sub TabStrip1_Click()

Dim i As Integer

i = TabStrip1.SelectedItem.Index

Frame3(i - 1).ZOrder

 If i = 2 Then
 conectWhat
 End If


End Sub

Private Sub Text1_DblClick()
Text1.Text = ""
End Sub

Private Sub Text1_GotFocus()
If AbiliteBloco(chaveMestra) Then
ConsultarTelefone
End If


End Sub

Private Sub Text1_LostFocus()
   AtLIberarManual
    Text6.Enabled = True
    Call Form_KeyDown(114, 0)
' Text6.SetFocus
End Sub

Private Sub Text10_DblClick()
Text10.Text = ""
End Sub

Private Sub Text11_Change()
Dim codCliente As Integer
 If Text11 <> "inicio" Then
 
 ConServer
 'SendKeys "%{TAB}"
 Form409.Text6.Text = Text11.Text
 
TransferirDadosDaConsulta (Text11.Text)
LiberarControles (2)
''cmdEntrega.Enabled = True

'consultarUtimosPedidos (Text11.Text)
EnabliteBloco
consultarUtimosPedidos (Text11.Text)
End If
End Sub

Private Sub Text3_DblClick()
Text3.Text = ""
End Sub

Private Sub Text4_Change()
'Dim fixoOUcelular As Integer
'If (Len(Text4.Text) = 10) Then
'Label4.Caption = "Telefone celular"
'telefoneTipo = 9
'Else
'Label4.Caption = "Telefone fixo"
'telefoneTipo = 3
'End If
COSULTAR = True
End Sub

Private Sub Text4_Click()
Form3.Show
COSULTAR = True
End Sub



Private Sub Text5_DblClick()
Text5.Text = ""
End Sub

Private Sub Text6_GotFocus()
Command13.Enabled = True
'buscainicialAiEntrega = False
End Sub
Private Sub Text6_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Then

If DataList1.Visible = True Then

'DataList1.SetFocus
End If
End If

End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If Text6.Text <> "" Then
Text6.Text = Trim(Text6.Text)
Call Command13_Click
If DataList1.Visible = True Then
'DataList1.SetFocus
End If
Else
DataList1.Visible = False
Frame4.Visible = False
'Text6.SetFocus
End If
End If
End Sub

Private Sub Text6_LostFocus()
If txtnumero.Enabled = True Then
'txtnumero.SetFocus
End If
End Sub

Private Sub Text8_Change()

'If buscainicialAiEntrega = False Then
'Text8.Text = Trim(Text8.Text)
'Text6.Text = Trim(Text6.Text)
'If Text8.Text <> "" And Text6.Text <> "" Or buscainicialAiEntrega = False Then
'AiEntrega
'Else
'Aibuscar
'End If
'End If
End Sub

Private Sub Text8_GotFocus()
'
'todo achando erro de no entrega

'buscainicialAiEntrega = False

'buscainicialAiEntrega = True
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'If Text8.Text <> "" And Text6.Text <> "" Then
'AiEntrega
'Else
'Aibuscar
'End If
'gerarlojaAutomatic
'End If

End Sub

Private Sub Text8_LostFocus()
'If Text8.Text <> "" And Text6.Text <> "" Then
'AiEntrega
'Else
'Aibuscar
'End If
gerarlojaAutomatic
End Sub

Private Sub Text9_DblClick()
Text9.Text = ""
End Sub

Private Sub Text9_LostFocus()
Text9.Text = Format(Text9.Text, ">")
End Sub

Private Sub Timer1_Timer()
If Text3.BackColor = &H80000005 And Text3 = "" Then
Text3.BackColor = &HC0FFC0
Else
Text3.BackColor = &H80000005
End If
End Sub

Private Sub Timer2_Timer()
If Text5.BackColor = &H80000005 And Text5 = "" Then
Text5.BackColor = &HC0FFC0
Else
Text5.BackColor = &H80000005
End If
End Sub


Private Sub tXTbAIRRO_Change()
If buscainicialAiEntrega = False Then
If tXTbAIRRO <> "" Then
AiEntrega
Else
Aibuscar
End If
End If
Text8.Text = tXTbAIRRO.Text
End Sub

Private Sub tXTcEP_DblClick()
tXTcEP.Text = ""
End Sub

Private Sub tXTcEP_KeyPress(KeyAscii As Integer)
'gerar evento
KeyAscii = Asc(UCase(Chr(KeyAscii)))

       ' Intercepta um código ASCII recebido e admite somente letras

          If InStr("AÃÁBCÇDEÉÊFGHIÍJKLMNOPQRSTUÚVWXYZ", Chr(KeyAscii)) = 0 Then
                
            If tXTcEP.SelStart = 5 Then tXTcEP.SelText = "-"
            If KeyAscii = 13 Then
            invisibiliteOstestAtenativos
           localizarCepsemposicionarMap
            End If
            If KeyAscii = 8 Or tXTcEP.Text = "" Then
            limpaPesquisacepBranco
            visibiliteOstestAtenativos
            AterarcoresdeboxparaRosa
            If KeyAscii = 8 And tXTcEP.Text = "" Then
            'Text6.SetFocus
            End If
            End If

                
                
                
         Else
         
         
         localizarCepsemposicionarMap
         'tXTcEP.Text = ""
         End If
    
        




            
gerarlojaAutomatic

End Sub

Private Sub TXTcIDADE_Change()
Text10.Text = TXTcIDADE.Text
End Sub

Private Sub txtEstado_Change()
Text9.Text = txtEstado.Text
End Sub




Private Sub txtnumero_Change()
If txtnumero.Text <> "" Then
Command3.Enabled = True
Else
Command3.Enabled = False
End If

End Sub

Private Sub txtnumero_DblClick()
txtnumero.Text = ""
End Sub

Private Sub txtnumero_LostFocus()
verificarSeApt (Trim(txtnumero))

'Text8.SetFocus

End Sub

Private Sub txtRua_Change()
Text6.Text = txtRua.Text
End Sub


Private Sub WebBrowser1_StatusTextChange(ByVal Text As String)
'Dim ftp As New ChilkatFtp2
'tXTcEP.Text = ""
End Sub

Public Sub localizarCep()
Adodc1.RecordSource = ""

Adodc1.CommandType = adCmdText

Adodc1.RecordSource = "SELECT * FROM CEP WHERE CEP ='" & tXTcEP.Text & "' "

Adodc1.Refresh
If Adodc1.Recordset.BOF = False Then
posicioneOmapa
'Set DataGrid1.DataSource = Adodc1
invisibiliteOstestAtenativos
If txtnumero.Enabled = True Then
'txtnumero.SetFocus
End If
Else

'tXTcEP.Text = ""

End If


End Sub

Public Sub posicioneOmapa()
On Error GoTo trata_erro
   WebBrowser1.Navigate Trim("http://maps.google.com/maps?q= Rua' " & txtRua.Text & " ',Bairro ' " & tXTbAIRRO.Text & " ',Cidade ' " & TXTcIDADE.Text & " '  ")
   Exit Sub
trata_erro:
  MsgBox Err.Description
'consultaEndereco.Append ("http://maps.google.com/maps?q=")


End Sub

Public Sub localizarPorRua(nome As String)
Adodc1.RecordSource = ""

Adodc1.CommandType = adCmdText

nome = Trim("%" + nome + "%")
'Adodc1.CommandType = adCmdText

'Adodc1.RecordSource = "SELECT * FROM `cadastro_de_usuarios_telemarketing` WHERE `nome` LIKE '" & nome & "'"

Adodc1.RecordSource = "SELECT * FROM CEP WHERE ENDERECO  LIKE '" & nome & "' "

Adodc1.Refresh
If Adodc1.Recordset.BOF = False Then

End If


End Sub

Public Sub limpaPesquisacepBranco()
Adodc1.RecordSource = ""

Adodc1.CommandType = adCmdText

Adodc1.RecordSource = "SELECT * FROM CEP WHERE CEP ='0' "

Adodc1.Refresh
If Adodc1.Recordset.BOF = False Then

End If

End Sub

Public Sub visibiliteOstestAtenativos()
Text6.Visible = True
Text8.Visible = True
Text10.Visible = True
Text9.Visible = True
EnvieManual = True



End Sub
Public Sub invisibiliteOstestAtenativos()
Text6.Visible = False
Text8.Visible = False
Text10.Visible = False
Text9.Visible = False
EnvieManual = False



End Sub


Public Sub AterarcoresdeboxparaBranco()
'&H00C0C0FF&
Text6.BackColor = &HFFFFFF
Text8.BackColor = &HFFFFFF
Text10.BackColor = &HFFFFFF
Text9.BackColor = &HFFFFFF

End Sub
Public Sub AterarcoresdeboxparaRosa()
'&H00C0C0FF&
Text6.BackColor = &HC0C0FF
Text8.BackColor = &HC0C0FF
Text10.BackColor = &HC0C0FF
Text9.BackColor = &HC0C0FF



   

End Sub

Public Sub EnabliteBloco()
Text4.Enabled = False
tXTcEP.Enabled = False
Text1.Enabled = False
Text6.Enabled = False
Text8.Enabled = False
Text9.Enabled = False
Text10.Enabled = False
txtnumero.Enabled = False
Text3.Enabled = False
Text5.Enabled = False
'DataCombo1.Enabled = False
Text2.Enabled = False




End Sub


Public Function AbiliteBloco(chave As Integer) As Boolean
Text4.Enabled = True
tXTcEP.Enabled = True
Text1.Enabled = True
Text6.Enabled = True
Text8.Enabled = True
Text9.Enabled = True
Text10.Enabled = True
txtnumero.Enabled = True
Text3.Enabled = True
Text5.Enabled = True
DataCombo1.Enabled = True
Text2.Enabled = True
If chave = 0 Then
AbiliteBloco = False
Else
AbiliteBloco = True
End If

End Function


Public Function ConsultarTelefone()
If COSULTAR = True Then
        LimparAtendimentopRINCIPAL
        ConServer
        Dim numeroTel As String
        
        
        Dim ENDERECOSnUM As Integer
        
        Dim sql As String
        Dim rs As New ADODB.Recordset
        Set rs = New ADODB.Recordset
        
        Set rs.ActiveConnection = con
        
        rs.CursorLocation = adUseClient
        'localizar telefone
        numeroTel = Trim("%" + Text4.Text + "%")
        
           sql = "SELECT COUNT(*) FROM `Cli_clientes` WHERE `telefone` LIKE '" & numeroTel & "' "
           rs.Open sql
           ENDERECOSnUM = rs.Fields("COUNT(*)").Value
           rs.Close
            
            If ENDERECOSnUM = 0 Then
             chaveEdicao = 2
             End If
           '
             If ENDERECOSnUM = 1 Then
             LiberarControles (3)
             End If
           
           
           If ENDERECOSnUM > 1 And Text12 <> "0" Then
           Form22.Show
           Form2.Hide
           'Form22.SetFocus
           
           sql = "SELECT `id`,`nome`,`endereco`,`numero`,`bairro`,`cidade`,`Apt`,`bloco` FROM `Cli_clientes` WHERE `telefone` LIKE '" & numeroTel & "' "
           rs.Open sql
           
           
           Set Form22.DataGrid1.DataSource = rs
           
           Else
          rs.Open sql
            If rs.BOF = False Then
            If Text12.Text = "0" And Text11.Text <> "" Then
            
             rs.Close
            sql = "SELECT * FROM `Cli_clientes` WHERE `id` = '" & Text11.Text & "' "
            'sql = "SELECT * FROM `Cli_clientes` WHERE `telefone` LIKE '3473-4959' "
            rs.Open sql
            Else
             rs.Close
            sql = "SELECT * FROM `Cli_clientes` WHERE `telefone` LIKE '" & numeroTel & "' "
            'sql = "SELECT * FROM `Cli_clientes` WHERE `telefone` LIKE '3473-4959' "
            rs.Open sql
        
        End If
            
                     If rs.BOF = False Then
                     EnabliteBloco
                     visibiliteOstestAtenativos
                     AterarcoresdeboxparaBranco
                     
                       tXTcEP.Text = rs.Fields("cep").Value
                       Text1.Text = rs.Fields("nome").Value
                       Text6.Text = rs.Fields("endereco").Value
                       txtnumero.Text = rs.Fields("numero").Value
                       Text3.Text = rs.Fields("Apt").Value
                       Text5.Text = rs.Fields("bloco").Value
                       Text2.Text = rs.Fields("referecia").Value
                       Text8.Text = rs.Fields("bairro").Value
                       Text9.Text = rs.Fields("uf").Value
                       Text10.Text = rs.Fields("cidade").Value
                      DataCombo1.Text = rs.Fields("loja").Value
                       'MsgBox "Cliente cadastrado", , "Cadastro ok"
                       resgateid
                      Else
                      LiberarControles (3)
                      invisibiliteOstestAtenativos
                      End If
           End If
           End If
'            If ENDERECOSnUM > 1 And Text12 <> "0" Then
'            Form22.Show
'
'           Form22.SetFocus
'           End If
        Set rs = Nothing
End If
COSULTAR = False


End Function


Public Sub TransferirDadosDaConsulta(buscarPor As Integer)



ConServer
Dim numeroTel As String

Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs.ActiveConnection = con

    
    
    sql = "SELECT * FROM `Cli_clientes` WHERE `id` = '" & buscarPor & "' "
    'sql = "SELECT * FROM `Cli_clientes` WHERE `telefone` LIKE '3473-4959' "
    rs.Open sql

    
             If rs.BOF = False Then
             EnabliteBloco
             visibiliteOstestAtenativos
             AterarcoresdeboxparaBranco
             
               tXTcEP.Text = rs.Fields("cep").Value
               Text1.Text = rs.Fields("nome").Value
               Text6.Text = rs.Fields("endereco").Value
               txtnumero.Text = rs.Fields("numero").Value
               Text3.Text = rs.Fields("Apt").Value
               Text5.Text = rs.Fields("bloco").Value
               Text2.Text = rs.Fields("referecia").Value
               Text8.Text = rs.Fields("bairro").Value
               Text9.Text = rs.Fields("uf").Value
               Text10.Text = rs.Fields("cidade").Value
              DataCombo1.Text = rs.Fields("loja").Value
               'MsgBox "Cliente cadastrado", , "Cadastro ok"
              
              End If
      
 Set rs = Nothing
End Sub

Public Sub REMOVespacosEmBranco()
              tXTcEP.Text = LTrim(tXTcEP.Text)
               Text1.Text = LTrim(Text1.Text)
               Text6.Text = LTrim(Text6.Text)
               txtnumero.Text = LTrim(txtnumero.Text)
               Text3.Text = LTrim(Text3.Text)
               Text5.Text = LTrim(Text5.Text)
               Text2.Text = LTrim(Text2.Text)
               Text8.Text = LTrim(Text8.Text)
               Text9.Text = LTrim(Text9.Text)
               Text10.Text = LTrim(Text10.Text)
              DataCombo1.Text = LTrim(DataCombo1.Text)

End Sub

Public Sub verificarSeApt(numero As String)
If tXTcEP <> "" And numero <> "" Then
ConServer

Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs.ActiveConnection = con

'localizar  localizar apartamento


   sql = "SELECT `Apt`,`bloco` FROM `Cli_clientes` WHERE `cep` LIKE '" & tXTcEP.Text & "' AND `numero` LIKE '" & txtnumero.Text & "' "
   rs.Open sql
    If rs.BOF = False Then
            If (rs.Fields("Apt").Value <> "" And rs.Fields("Apt").Value <> " " And rs.Fields("Apt").Value <> "  ") Then
            Timer1.Interval = 500
            End If
             If (rs.Fields("bloco").Value <> "" And rs.Fields("bloco").Value <> " " And rs.Fields("bloco").Value <> "  ") Then
              Timer2.Interval = 500
            End If
            Else
Timer1.Interval = 0
Timer2.Interval = 0
Text3.BackColor = &H80000005
Text5.BackColor = &H80000005
End If
      
   rs.Close
   
  
    

 Set rs = Nothing
Else
Timer1.Interval = 0
Timer2.Interval = 0
Text3.BackColor = &H80000005
Text5.BackColor = &H80000005
End If
End Sub

Public Sub localizarCepsemposicionarMap()
Adodc1.RecordSource = ""

Adodc1.CommandType = adCmdText

Adodc1.RecordSource = "SELECT * FROM CEP WHERE CEP ='" & tXTcEP.Text & "' "

Adodc1.Refresh
If Adodc1.Recordset.BOF = False Then
'posicioneOmapa
'Set DataGrid1.DataSource = Adodc1
invisibiliteOstestAtenativos
'txtnumero.SetFocus
Else

'tXTcEP.Text = ""

End If

End Sub

Public Sub LiberarControles(Optiones As Integer)

Select Case Optiones
      Case 0 ' fechar todos  todos
            CmdConsultar.Enabled = False
            cmdEditar.Enabled = False
            cmdNovo.Enabled = False
            CmdLimpar.Enabled = False
            cmdSalvar.Enabled = False
     Case 1 ' libera para editar ou novo limpar *
             cmdNovo.Enabled = True
            'CmdLimpar.Enabled = True
            cmdEditar.Enabled = True
            'cmdsalvar.Enabled=False
     
     
     Case 2 ' fechar todos  todos * consulta retorna
     
            CmdConsultar.Enabled = True
            cmdEditar.Enabled = True
            cmdNovo.Enabled = True
            'CmdLimpar.Enabled = True
            cmdSalvar.Enabled = False
        Case 3 ' situação novo cliente
     
            'CmdConsultar.Enabled = True
            cmdEditar.Enabled = True
            cmdNovo.Enabled = True
            CmdLimpar.Enabled = True
            cmdSalvar.Enabled = False
        
        Case 4 ' situação novo cliente
     
            'CmdConsultar.Enabled = True
            cmdEditar.Enabled = True
            cmdNovo.Enabled = True
            CmdLimpar.Enabled = True
            cmdSalvar.Enabled = False
            
End Select

End Sub

Public Sub CaixaAlta()
               tXTcEP.Text = Format(LTrim(tXTcEP.Text), ">")
               Text1.Text = Format(LTrim(Text1.Text), ">")
               Text6.Text = Format(LTrim(Text6.Text), ">")
               txtnumero.Text = Format(LTrim(txtnumero.Text), ">")
               Text3.Text = Format(LTrim(Text3.Text), ">")
               Text5.Text = Format(LTrim(Text5.Text), ">")
               Text2.Text = Format(LTrim(Text2.Text), ">")
               Text8.Text = Format(LTrim(Text8.Text), ">")
               Text9.Text = Format(LTrim(Text9.Text), ">")
               Text10.Text = Format(LTrim(Text10.Text), ">")
              DataCombo1.Text = Format(LTrim(DataCombo1.Text), ">")
              'bloco neutro
               tXTbAIRRO.Text = Format(LTrim(tXTbAIRRO.Text), ">")
               txtRua.Text = Format(LTrim(txtRua.Text), ">")
               TXTcIDADE.Text = Format(LTrim(TXTcIDADE.Text), ">")
               txtEstado.Text = Format(LTrim(txtEstado.Text), ">")


End Sub
Public Sub LimparAtendimentopRINCIPAL()

                    tXTcEP.Text = ""
               Text1.Text = ""
               Text6.Text = ""
               txtnumero.Text = ""
               Text3.Text = ""
               Text5.Text = ""
               Text2.Text = ""
               Text8.Text = ""
               Text9.Text = ""
               Text10.Text = ""
              DataCombo1.Text = ""
              'bloco neutro
             visibiliteOstestAtenativos
             
             Text4.Enabled = True
tXTcEP.Enabled = True
Text1.Enabled = True
Text6.Enabled = True
Text8.Enabled = True
Text9.Enabled = True
Text10.Enabled = True
txtnumero.Enabled = True
Text3.Enabled = True
Text5.Enabled = True
DataCombo1.Enabled = True
Text2.Enabled = True
             
             
             Text1.SetFocus
             

End Sub
Public Sub LimparAtendimento()

                    tXTcEP.Text = ""
               Text1.Text = ""
               Text6.Text = ""
               txtnumero.Text = ""
               Text3.Text = ""
               Text5.Text = ""
               Text2.Text = ""
               Text8.Text = ""
               Text9.Text = ""
               Text10.Text = ""
              DataCombo1.Text = ""
              'bloco neutro
             visibiliteOstestAtenativos
             'tXTcEP.SetFocus

End Sub

Public Sub cauculafrete(valordofrete As Double)
'valor do frete para outro formulario
'Adodc2.RecordSource = ""
'
'Adodc2.CommandType = adCmdText
'
'Adodc2.RecordSource = "SELECT `frete` FROM `cadastro_da_empresa` WHERE `bairro` LIKE '" & Trim(DataCombo1.Text) & "'"
'
'Adodc2.Refresh
'If Adodc2.Recordset.BOF = False Then
'
'Form401.Show
'Form401.Text23.Text = 0
'Form401.Text12.Text = Format(valordofrete, "currency")
''Adodc2.Recordset.Close
'Adodc2.RecordSource = ""
'
'Adodc2.CommandType = adCmdText
'
'Adodc2.RecordSource = "SELECT * FROM `cadastro_da_empresa` ORDER BY `cadastro_da_empresa`.`bairro` ASC"
''Adodc2.Recordset.Open
'Adodc2.Refresh
'
'End If

End Sub

Public Sub pesquisarFreteassinado()

ConServer
Dim bEntrega As String
Dim bloja As String
Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset
Set rs.ActiveConnection = con

If DataCombo1.Text <> "" And DataCombo1.Text <> " " And DataCombo1.Text <> "  " And DataCombo1.Text <> "Selecione a Loja!" Then
                If Text8.Text <> "" Then
                bEntrega = Trim(LTrim(Text8.Text))
                Else
                bEntrega = Trim(LTrim(tXTbAIRRO))
                End If
                If bdEntreganull = "" Then
                bEntrega = Trim(bdEntreganull)
                End If
                bloja = Trim(LTrim(DataCombo1.Text))
   
                sql = "SELECT * FROM `re_fretebairroloja` WHERE `bairro` LIKE '" & bEntrega & "' AND `loja` LIKE '" & bloja & "'"
                 rs.Open sql
                  If rs.BOF = True Then
                  rs.Close
                  If Text8.Text = "" Or bdEntreganull = "" Then
                  
                  criefretebuscar
                  bEntrega = ""
                  Form500.Hide
                  End If
                 sql = "SELECT * FROM `re_fretebairroloja` WHERE `bairro` LIKE '" & Format(bEntrega, ">") & "' AND `loja` LIKE '" & Format(bloja, ">") & "'"
                 rs.Open sql
                If rs.BOF = False Then
                End If
                 End If
                 Form500.ProgressBar1.Value = 80
                    If rs.BOF = False Then
                        Animation1.Open Text15.Text

                        Animation1.AutoPlay = True
                                   
                                    Form401.Show
                                    Form401.Text23.Text = 0
                                     If bdEntreganull = "" Then
                                     Form401.Text12.Text = Format("0", "currency")
                                                                         Form401.LabelFRETE = Format("0", "currency")
                                     Else
                                    Form401.Text12.Text = Format(rs.Fields("valor").Value, "currency")
                                                                        Form401.LabelFRETE = Format(rs.Fields("valor").Value, "currency")
                                    End If

                                    Command6.Enabled = True
                                    Command2.Enabled = False
                                    Animation1.Close

    Else
    Form23.Visible = True
    Form23.Text3.SetFocus
    
    Form23.DataCombo2 = bEntrega
    Form23.DataCombo1 = bloja
    'Form23.Text3.SetFocus
    If Form23.DataCombo1 = "" Then
    pesquisarFreteassinado
     End If
     End If

    


                                    Else
                                    Animation1.AutoPlay = True
                                    MsgBox "Inclua a loja de destino", , "Loja não foi selecionada!"
                                    If DataCombo1.Enabled = True Then
                                    'DataCombo1.SetFocus
                                    Else
                                    MsgBox "Clique no icone para editar assim poderá alterar a loja de destino", , "Loja não foi selecionada!"
                                    'cmdEditar.SetFocus
                                    End If
                                    End If

Form500.ProgressBar1.Value = 100
obtervalordeFrete
 Set rs = Nothing
'INSERT INTO `robofi61_order_taker`.`re_fretebairroloja` (`bairro`, `loja`, `valor`, `user`, `created`, `modified`) VALUES ('bteste', 'lojatest', '15.50', 'testeuse', '00:00:00 00/00/0000', '00:00:00 00/00/0000');
'UPDATE `robofi61_order_taker`.`re_fretebairroloja` SET `bairro` = 'bteste1', `loja` = 'lojatest1', `valor` = '15.51', `user` = 'testeuse1', `modified` = '0000-00-00 00:00:01' WHERE ;

End Sub


Public Sub resgateid()
If Text11.Text = "inicio" Then
            ConServer
            Dim numeroTel As String
            
            Dim id As Integer
            
            Dim sql As String
            Dim rs As New ADODB.Recordset
            Set rs = New ADODB.Recordset
            
            Set rs.ActiveConnection = con
            
            rs.CursorLocation = adUseClient
            'localizar telefone
            numeroTel = Trim("%" + Text4.Text + "%")
            
               sql = "SELECT `id` FROM `Cli_clientes` WHERE `telefone` LIKE '" & numeroTel & "' "
               rs.Open sql
               If rs.BOF = False Then
              Text11.Text = rs.Fields("id").Value
              End If
              rs.Close
             
            Set rs = Nothing
 End If
End Sub

Public Sub consultarUtimosPedidos(fkCliente As Integer)
ConServer
'
            Dim DataPedido As Date
            Dim dataDoSistema As Date
            Dim intervaloDeDatas As Integer

            Dim HorasPedido As String
            Dim horasDoSistema As String
            Dim intervaloDehoras As Integer

            Dim sql As String
            Dim rs As New ADODB.Recordset
            
            
            
            Set rs = New ADODB.Recordset

            Set rs.ActiveConnection = con

            rs.CursorLocation = adUseClient
            'localizar telefone


               sql = "SELECT `datahora`,`numPedido`,`total` FROM `at_Cupon` WHERE `fk_Cliente` = '" & fkCliente & "'ORDER BY `id` DESC"
               rs.Open sql
               If rs.BOF = False Then
               Frame1.Visible = True


               Set DataGrid1.DataSource = rs
               'DataPedido = Format(rs.Fields("datahora").Value, "dd/mm/yyyy")
               'dataDoSistema = Format(Now, "dd/mm/yyyy")
               'intervaloDeDatas = DateDiff("d", DataPedido, dataDoSistema)
                '     If intervaloDeDatas = 0 Then
                     'SendKeys "%{F2}"
                 '    MsgBox "Esse cliente fez um pedido hoje!", vbInformation, "Possível alteração identificada "
                       'SendKeys "%{TAB}"
                  '   End If
               End If
             ' rs.Close

            Set rs = Nothing
 
End Sub


Public Sub consultaritensUtimosPedidos(fkItensdoPedido As Integer)
ConServer
   
            
            Dim sql As String
            Dim rs As New ADODB.Recordset
            Set rs = New ADODB.Recordset
            
            Set rs.ActiveConnection = con
            
            rs.CursorLocation = adUseClient
            'localizar telefone
           
            
               sql = "SELECT * FROM `at_itens` WHERE `fk_pedido` =  '" & fkItensdoPedido & "'ORDER BY `id` DESC"
               rs.Open sql
               If rs.BOF = False Then
               Frame1.Visible = True
               Set DataGrid1.DataSource = rs
              End If
             ' rs.Close
             
            Set rs = Nothing
 
End Sub


Public Sub CancelamentoPedido()

ConServer

Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs.ActiveConnection = con
rs.CursorLocation = adUseClient






sql = "UPDATE `at_contadorDePedidos` SET `contador` = '0', `situacaoImpressao` = '3'  WHERE `at_contadorDePedidos`.`id` =  '" & numPedido & "'"
                            rs.Open sql
'rs.Close
sql = "DELETE FROM `at_itens` WHERE `fk_pedido` = '" & numPedido & "' ORDER BY `id` DESC"
 rs.Open sql
 sql = "DELETE  FROM `at_Cupon` WHERE `numPedido` = '" & numPedido & "' ORDER BY `numPedido` ASC"
  rs.Open sql
sql = " DELETE FROM `Pagamento` WHERE `NUmPedido` ='" & numPedido & "' ORDER BY `NUmPedido` ASC"
 rs.Open sql
sql = "UPDATE `at_contadorDePedidos` SET `contador` = '0' WHERE `at_contadorDePedidos`.`id` =  '" & numPedido & "'"
                            rs.Open sql



 Set rs = Nothing
'End If
'comander6
'imprimirCupon

Form401.Text26.Text = 1
Form404.Text9.Text = 1
Unload Form404
Unload Form401
Unload Form402
Form2.Show
End Sub

Public Sub AlteracaoPedido()

ConServer

Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs.ActiveConnection = con
rs.CursorLocation = adUseClient






sql = "UPDATE `at_contadorDePedidos` SET `contador` = '0', `situacaoImpressao` = '3'  WHERE `at_contadorDePedidos`.`id` =  '" & numPedido & "'"
                            rs.Open sql
'rs.Close
'sql = "DELETE FROM `at_itens` WHERE `fk_pedido` = '" & numPedido & "' ORDER BY `id` DESC"
' rs.Open sql
' sql = "DELETE  FROM `at_Cupon` WHERE `numPedido` = '" & numPedido & "' ORDER BY `numPedido` ASC"
'  rs.Open sql
'sql = " DELETE FROM `Pagamento` WHERE `NUmPedido` ='" & numPedido & "' ORDER BY `NUmPedido` ASC"
' rs.Open sql
sql = "UPDATE `at_contadorDePedidos` SET `contador` = '0' WHERE `at_contadorDePedidos`.`id` =  '" & numPedido & "'"
                            rs.Open sql



 Set rs = Nothing
'End If
'comander6
'imprimirCupon

Form401.Text26.Text = 1
Form404.Text9.Text = 1
'Unload Form404
'Unload Form401
'Unload Form402
Form2.Show
End Sub



Public Sub verificarsepossuiaFrete()
Form404.Adodc2.RecordSource = ""

Form404.Adodc2.CommandType = adCmdText

Form404.Adodc2.RecordSource = "SELECT `frete` FROM `at_frete` WHERE `fk_numPedido` ='" & numPedido & "'"

Form404.Adodc2.Refresh
Form500.ProgressBar1.Value = 60
If Form404.Adodc2.Recordset.BOF = False Then
Form407.Text7.Text = Replace(Form404.Adodc2.Recordset.Fields("frete").Value, ",", ".")
Else
ConServer

Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs.ActiveConnection = con
rs.CursorLocation = adUseClient


sql = "INSERT INTO `at_frete` (`id`, `frete`, `fk_numPedido`) VALUES (NULL, '0', '" & numPedido & "') "
                            rs.Open sql


 Set rs = Nothing

End If
End Sub

Public Sub gerarlojaAutomatic()
Dim Bairro As String
Dim loja As String
Dim NumRegistros As Integer
If Text8.Text <> "" Or tXTbAIRRO.Text <> "" Then
If Text8.Text <> "" Then
Bairro = Trim("%" + Text8.Text + "%")
End If
If tXTbAIRRO.Text <> "" Then
Bairro = Trim("%" + tXTbAIRRO.Text + "%")
End If



ConServer

Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs.ActiveConnection = con
rs.CursorLocation = adUseClient

'counte os registros
sql = "SELECT COUNT(DISTINCT`loja`) FROM `re_fretebairroloja` WHERE `bairro` LIKE '" & Bairro & "' ORDER BY `re_fretebairroloja`.`id` DESC"
'sql = "SELECT COUNT(DISTINCT`loja`) FROM `re_fretebairroloja` WHERE `bairro` LIKE '%pindorama%'"
rs.Open sql
NumRegistros = rs.Fields("COUNT(DISTINCT`loja`)").Value
If NumRegistros <> 0 Then
        If NumRegistros > 1 Then
        DataCombo1.Text = "Selecione a Loja!"
        
           rs.Close
           sql = "SELECT DISTINCT (`loja`) FROM `re_fretebairroloja` WHERE `bairro` LIKE '%" & Bairro & "%'"
           rs.Open sql
            If rs.BOF = False Then
            Set DataCombo1.RowSource = rs
        
            DataCombo1.ListField = "loja"
            End If
        Else
        rs.Close
           sql = "SELECT DISTINCT (`loja`) FROM `re_fretebairroloja` WHERE `bairro` LIKE '" & Bairro & "' ORDER BY `re_fretebairroloja`.`id` DESC"
          ' sql = "SELECT DISTINCT (`loja`) FROM `re_fretebairroloja` WHERE `bairro` LIKE '%pindorama%'"
           rs.Open sql
            If rs.BOF = False Then
            Set DataCombo1.RowSource = rs
        
            DataCombo1.ListField = "loja"
            DataCombo1.Text = rs.Fields("loja").Value
           End If
      End If
     
      
Else
 
         rs.Close
           sql = "SELECT * FROM `cadastro_da_empresa` ORDER BY `cadastro_da_empresa`.`bairro` ASC"
           rs.Open sql
            If rs.BOF = False Then
            Set DataCombo1.RowSource = rs
        
            DataCombo1.ListField = "bairro"
            'DataCombo1.Text = rs.Fields("loja").Value
           End If
           DataCombo1.Text = " "
End If
    

  ' rs.Close
   
  
    


End If




Set rs = Nothing
End Sub

Public Sub obtervalordeFrete()
Dim loja As String
Dim Bairro As String
Bairro = Trim(Text8.Text)
'Bairro = tXTbAIRRO.Text

loja = Trim(DataCombo1.Text)
ConServer


Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs.ActiveConnection = con


rs.CursorLocation = adUseClient
  sql = "SELECT `valor` FROM `re_fretebairroloja` WHERE `bairro` LIKE '" & Bairro & "' AND `loja` LIKE  '%" & loja & "%'  LIMIT 1"
  rs.Open sql
  If rs.BOF = False Then
  Label13.Caption = rs.Fields("valor").Value
  End If
 Set rs = Nothing
Label13.Visible = True
Label14.Visible = True
  
  
End Sub

Public Sub criefretebuscar()


ConServer

Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset
Set rs.ActiveConnection = con
   'Label1.Caption = Format(Now, "yyyy/mm/dd  hh:mm:ss ")
   'Label1.Caption = Format(Label1.Caption, "DataValue")
    sql = "INSERT INTO `robofi61_order_taker`.`re_fretebairroloja` (`bairro`, `loja`, `valor`, `user`, `created`, `modified`) VALUES ('', '" & Format(LTrim(DataCombo1.Text), ">") & "', '0', 'sistema', '" & Format(Now, "yyyy/mm/dd  hh:mm:ss ") & "', '" & Format(Now, "yyyy/mm/dd  hh:mm:ss ") & "')"
    rs.Open sql
    
 Set rs = Nothing




'INSERT INTO `robofi61_order_taker`.`re_fretebairroloja` (`bairro`, `loja`, `valor`, `user`, `created`, `modified`) VALUES ('bteste', 'lojatest', '15.50', 'testeuse', '00:00:00 00/00/0000', '00:00:00 00/00/0000');
'UPDATE `robofi61_order_taker`.`re_fretebairroloja` SET `bairro` = 'bteste1', `loja` = 'lojatest1', `valor` = '15.51', `user` = 'testeuse1', `modified` = '0000-00-00 00:00:01' WHERE ;

End Sub

Public Sub AiEntrega()
 'Text8.Text = varText8
'Text6.Text = varText6
CmdLojas.Enabled = True
'cmdEntrega.Enabled = True
bdEntreganull = "entregas"
cmdEntrega.Picture = Picture2.Image
cmdEntrega.Caption = "Entrega"
TROCAimagem = False
tXTcEP.Text = Trim(tXTcEP.Text)
Text8.Text = Trim(Text8.Text)
If Text8.Text = "" And Text11.Text <> "inicio" Then
'SALVANDO AUTOMATIC

Call cmdEditar_Click
tXTcEP.Enabled = True
'tXTcEP.SetFocus
End If

End Sub

Public Sub Aibuscar()
'cmdEntrega.Enabled = false
'varText8 = Text8.Text
'varText6 = Text6.Text
'Text8.Text = ""
'Text6.Text = ""
CmdLojas.Enabled = True
 
cmdEntrega.Picture = Picture1.Image
cmdEntrega.Caption = "Buscar"
DataCombo1.Enabled = True '
TROCAimagem = True
bdEntreganull = ""
AiLiberarLojas
End Sub
Public Sub AiLiberarLojas()


ConServer

Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs.ActiveConnection = con
rs.CursorLocation = adUseClient

           sql = "SELECT * FROM `cadastro_da_empresa` ORDER BY `cadastro_da_empresa`.`bairro` ASC"
           rs.Open sql
            If rs.BOF = False Then
            Set DataCombo1.RowSource = rs
        
            DataCombo1.ListField = "bairro"
            'DataCombo1.Text = rs.Fields("loja").Value
           End If
'           DataCombo1.Text = " "

    

 




 Set rs = Nothing
End Sub

Public Sub AtLIberarManual()
             EnvieManual = True
            CmdLojas.Enabled = True
             tXTcEP.Text = ""
            limpaPesquisacepBranco
            visibiliteOstestAtenativos
            AterarcoresdeboxparaRosa
           
            'Text6.SetFocus
           
            
End Sub


Public Sub conectWhat()
Dim command As String

command = "C:\Order_Taker\whatsapp.bat"
Shell "cmd.exe /c " & command
'https://www.google.com/maps/@-19.904281,-44.022235,16z
On Error GoTo trata_erro
  WebBrowser2.Navigate Trim("https://web.whatsapp.com/")
   Exit Sub
trata_erro:
   MsgBox Err.Description
'consultaEndereco.Append ("http://maps.google.com/maps?q=")
End Sub

Public Function localizarendereco(rua As String)

rua = Trim("%" + rua + "%")
Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs.ActiveConnection = conl
rs.CursorLocation = adUseClient

           sql = "SELECT * FROM CEP WHERE `ENDERECO` LIKE '" & rua & "' "
           rs.Open sql
            If rs.BOF = False Then
            Set DataList1.RowSource = rs
             
            DataList1.ListField = "ENDERECO"
            'DataCombo1.Text = rs.Fields("loja").Value
           End If
           DataList1.Text = " "

    

 




Set rs = Nothing
End Function


Public Function Dinamic(rua As String)
Dim RUAL As String
ConServerloc
RUAL = rua
rua = Trim("%" + rua + "%")
Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs.ActiveConnection = conl
rs.CursorLocation = adUseClient

           sql = "SELECT * FROM CEP WHERE `ENDERECO` LIKE '" & rua & "' "
           rs.Open sql
            If rs.BOF = False Then
            Text6.Text = RUAL
            Set DataList1.RowSource = rs
            Set Text8.DataSource = rs
             Set Text9.DataSource = rs
             Set Text10.DataSource = rs
              Set tXTcEP.DataSource = rs
             
            DataList1.ListField = "ENDERECO"
            Text8.DataField = "BAIRRO"
            Text9.DataField = "UF"
            Text10.DataField = "CIDADE"
            tXTcEP.DataField = "CEP"
            'DataCombo1.Text = rs.Fields("loja").Value
            DataList1.Visible = False
            Frame4.Visible = False
           End If
           DataList1.Text = " "

    

 




 Set rs = Nothing
End Function

Public Sub gireoDATACOMBOlojadopedidoeincluaOPERADOR()



ConServer

Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs.ActiveConnection = con
rs.CursorLocation = adUseClient

           sql = "SELECT `intPedido`  FROM `at_contadorDePedidos` WHERE `id` = '" & DataGrid1.Columns(1).Value & "' "
           rs.Open sql
            If rs.BOF = False Then
            

            DataCombo1.Text = rs.Fields("intPedido").Value
          
           End If
           


End Sub

Public Sub ReverifiqueEntregaouBuscar()
If buscainicialAiEntrega = False Then
Text8.Text = Trim(Text8.Text)
Text6.Text = Trim(Text6.Text)
If Text8.Text <> "" And Text6.Text <> "" And buscainicialAiEntrega = False Then
AiEntrega
Else
Aibuscar
End If
End If
End Sub

Public Sub tragaAlojaDopedido()

End Sub
