VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form4441 
   Caption         =   "Form14"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   21030
   Enabled         =   0   'False
   LinkTopic       =   "Form14"
   ScaleHeight     =   9180
   ScaleWidth      =   21030
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text15 
      DataField       =   "CardapioOficial"
      DataSource      =   "Adodc2"
      Height          =   375
      Left            =   10680
      TabIndex        =   26
      Text            =   "Text15"
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox Text14 
      DataField       =   "nomeOriginaldoBtn"
      DataSource      =   "Adodc2"
      Height          =   375
      Left            =   8280
      TabIndex        =   25
      Text            =   "Text14"
      Top             =   600
      Width           =   2055
   End
   Begin VB.TextBox Text13 
      Height          =   495
      Left            =   10680
      TabIndex        =   24
      Text            =   "Text13"
      Top             =   0
      Width           =   735
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "Form4441.frx":0000
      Height          =   1455
      Left            =   6240
      TabIndex        =   22
      Top             =   6960
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   2566
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
   Begin VB.TextBox Text12 
      DataField       =   "IndiceProtocoloCardapio"
      DataSource      =   "Adodc2"
      Height          =   375
      Left            =   10680
      TabIndex        =   20
      Text            =   "Text12"
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   375
      Left            =   4080
      TabIndex        =   19
      Top             =   7440
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form4441.frx":0015
      Height          =   5175
      Left            =   12360
      TabIndex        =   18
      Top             =   360
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   9128
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
      Caption         =   "ligado ao adodc3 dever apresentar os tipos de acordo com o titulo selecionado"
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
   Begin VB.TextBox Text11 
      Height          =   855
      Left            =   480
      TabIndex        =   16
      Text            =   "Text11"
      Top             =   5640
      Width           =   1815
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   8040
      Top             =   5640
      Width           =   2895
      _ExtentX        =   5106
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
      RecordSource    =   "configuracao_tela_pedido"
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
   Begin VB.TextBox Text1 
      DataField       =   "imagemCaminho"
      DataSource      =   "Adodc2"
      Height          =   285
      Left            =   10560
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   1680
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3120
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text10 
      DataField       =   "fonte5"
      DataSource      =   "Adodc2"
      Height          =   285
      Left            =   8280
      TabIndex        =   12
      Text            =   "Text10"
      Top             =   5160
      Width           =   2175
   End
   Begin VB.TextBox Text9 
      DataField       =   "fonte4"
      DataSource      =   "Adodc2"
      Height          =   375
      Left            =   8280
      TabIndex        =   11
      Text            =   "Text9"
      Top             =   4560
      Width           =   2055
   End
   Begin VB.TextBox Text8 
      DataField       =   "fonte3"
      DataSource      =   "Adodc2"
      Height          =   375
      Left            =   8280
      TabIndex        =   10
      Text            =   "Text8"
      Top             =   3960
      Width           =   2055
   End
   Begin VB.TextBox Text7 
      DataField       =   "fonte2"
      DataSource      =   "Adodc2"
      Height          =   375
      Left            =   8280
      TabIndex        =   9
      Text            =   "Text7"
      Top             =   3360
      Width           =   1935
   End
   Begin VB.TextBox Text6 
      DataField       =   "fonte1"
      DataSource      =   "Adodc2"
      Height          =   375
      Left            =   8280
      TabIndex        =   8
      Text            =   "Text6"
      Top             =   2880
      Width           =   1935
   End
   Begin VB.TextBox Text5 
      DataField       =   "fonte0"
      DataSource      =   "Adodc2"
      Height          =   375
      Left            =   8280
      TabIndex        =   7
      Text            =   "Text5"
      Top             =   2280
      Width           =   1935
   End
   Begin VB.TextBox Text4 
      DataField       =   "texto"
      DataSource      =   "Adodc2"
      Height          =   375
      Left            =   8280
      TabIndex        =   6
      Text            =   "Text4"
      Top             =   1680
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      DataField       =   "Color"
      DataSource      =   "Adodc2"
      Height          =   375
      Left            =   8280
      TabIndex        =   5
      Text            =   "Text3"
      Top             =   1080
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   405
      Left            =   8280
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   0
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "Personalize"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   4200
      TabIndex        =   3
      Top             =   600
      Width           =   3495
      Begin MSAdodcLib.Adodc Adodc3 
         Height          =   375
         Left            =   480
         Top             =   120
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   4471
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
         RecordSource    =   "Cardapio"
         Caption         =   "Adodc3"
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
      Begin VB.CommandButton Command2 
         Caption         =   "Cores"
         Height          =   495
         Left            =   360
         TabIndex        =   17
         Top             =   1320
         Width           =   2895
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "Form4441.frx":002A
         Height          =   420
         Left            =   360
         TabIndex        =   14
         Top             =   600
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   741
         _Version        =   393216
         ListField       =   "tipo"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Imagem"
         Height          =   615
         Left            =   360
         TabIndex        =   13
         Top             =   2760
         Width           =   2895
      End
      Begin VB.CommandButton cmdSalvar 
         Caption         =   "Salvar"
         Height          =   615
         Left            =   1440
         Picture         =   "Form4441.frx":003F
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   3720
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Fonte"
         Height          =   495
         Left            =   360
         TabIndex        =   0
         Top             =   2040
         Width           =   2895
      End
      Begin VB.Label Label2 
         Caption         =   "Texto"
         Height          =   375
         Left            =   1560
         TabIndex        =   23
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Personalize"
      Height          =   2055
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1800
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Protocolo cardapio"
      Height          =   375
      Left            =   10680
      TabIndex        =   21
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000006&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Left            =   1080
      Top             =   1920
      Width           =   2655
   End
End
Attribute VB_Name = "Form4441"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


