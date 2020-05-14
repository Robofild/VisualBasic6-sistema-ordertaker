VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.ocx"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "comct232.ocx"
Begin VB.Form Form401 
   Caption         =   "Atendimento em andamento"
   ClientHeight    =   10260
   ClientLeft      =   1815
   ClientTop       =   1980
   ClientWidth     =   17940
   Icon            =   "NEWNEWaTENDIMENTO.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form9"
   Moveable        =   0   'False
   ScaleHeight     =   10260
   ScaleWidth      =   17940
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command13 
      Caption         =   "Command13"
      Height          =   615
      Left            =   17880
      TabIndex        =   75
      Top             =   960
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text33 
      Height          =   285
      Left            =   6240
      TabIndex        =   74
      Text            =   "?"
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Command12"
      Height          =   495
      Left            =   18840
      TabIndex        =   72
      Top             =   4440
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text32 
      Height          =   375
      Left            =   20160
      TabIndex        =   71
      Text            =   "0"
      Top             =   3720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text31 
      Height          =   375
      Left            =   20160
      TabIndex        =   67
      Text            =   "0"
      Top             =   2880
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Command11"
      Height          =   615
      Left            =   17040
      TabIndex        =   66
      Top             =   6120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text30 
      Height          =   285
      Left            =   18240
      TabIndex        =   64
      Text            =   "0"
      Top             =   2760
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Command10"
      Height          =   495
      Left            =   17040
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   4680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text29 
      Enabled         =   0   'False
      Height          =   375
      Left            =   17040
      TabIndex        =   60
      Text            =   "0"
      Top             =   1200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text28 
      Height          =   615
      Left            =   14760
      TabIndex        =   58
      Top             =   8880
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox Text27 
      Height          =   375
      Left            =   3720
      TabIndex        =   57
      Text            =   "Text27"
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text26 
      Height          =   285
      Left            =   3000
      TabIndex        =   56
      Text            =   "1"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text25 
      Height          =   375
      Left            =   17880
      TabIndex        =   55
      Text            =   "Text25"
      Top             =   8400
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Timer Timer2 
      Left            =   20640
      Top             =   1920
   End
   Begin VB.TextBox Text24 
      Height          =   495
      Left            =   18120
      TabIndex        =   54
      Top             =   3120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text23 
      Enabled         =   0   'False
      Height          =   375
      Left            =   17040
      TabIndex        =   53
      Text            =   "Text23"
      Top             =   240
      Visible         =   0   'False
      Width           =   735
   End
   Begin ComCtl2.Animation Animation1 
      Height          =   855
      Left            =   16680
      TabIndex        =   27
      Top             =   2760
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1508
      _Version        =   327681
      FullWidth       =   81
      FullHeight      =   57
   End
   Begin VB.TextBox Text22 
      Height          =   375
      Left            =   16920
      TabIndex        =   50
      Text            =   "Text22"
      Top             =   1920
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.TextBox Text21 
      Height          =   375
      Left            =   19200
      TabIndex        =   49
      Text            =   "Text21"
      Top             =   6960
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text20 
      Height          =   375
      Left            =   19200
      TabIndex        =   48
      Text            =   "Text20"
      Top             =   6000
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text19 
      Height          =   375
      Left            =   18840
      TabIndex        =   47
      Text            =   "Text19"
      Top             =   3720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text18 
      Height          =   495
      Left            =   18960
      TabIndex        =   46
      Text            =   "Text18"
      Top             =   1080
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Text17 
      Height          =   495
      Left            =   18960
      TabIndex        =   45
      Text            =   "Text17"
      Top             =   240
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Command9"
      Height          =   495
      Left            =   17880
      TabIndex        =   43
      Top             =   480
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text15 
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
      Left            =   13320
      TabIndex        =   41
      Text            =   "1"
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Command8"
      Height          =   615
      Left            =   18360
      TabIndex        =   40
      Top             =   5400
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      Caption         =   "&F4            Ver o pedido"
      Enabled         =   0   'False
      Height          =   735
      Left            =   14880
      Picture         =   "NEWNEWaTENDIMENTO.frx":0A02
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   7800
      Width           =   1815
   End
   Begin VB.TextBox Text14 
      BackColor       =   &H00C0FFC0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   14760
      TabIndex        =   37
      Top             =   6960
      Width           =   1935
   End
   Begin VB.TextBox Text13 
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   360
      TabIndex        =   36
      Top             =   9600
      Width           =   14055
   End
   Begin VB.TextBox Text12 
      BackColor       =   &H00C0FFC0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   14760
      TabIndex        =   34
      Top             =   5880
      Width           =   1935
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&F3            Observações"
      Enabled         =   0   'False
      Height          =   735
      Left            =   14760
      Picture         =   "NEWNEWaTENDIMENTO.frx":1404
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   4680
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&F2             Acressímos"
      Enabled         =   0   'False
      Height          =   735
      Left            =   14760
      Picture         =   "NEWNEWaTENDIMENTO.frx":1E06
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   3720
      Width           =   1815
   End
   Begin VB.PictureBox Picture2 
      Height          =   375
      Left            =   15000
      Picture         =   "NEWNEWaTENDIMENTO.frx":2908
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   31
      Top             =   7800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text10 
      Height          =   615
      Left            =   17160
      TabIndex        =   30
      Text            =   "Text10"
      Top             =   7680
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   615
      Left            =   17040
      TabIndex        =   29
      Top             =   5400
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   495
      Left            =   18720
      TabIndex        =   28
      Top             =   7920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text11 
      Height          =   375
      Left            =   14640
      TabIndex        =   26
      Text            =   "C:\Order_Taker\FINDFILE.AVI"
      Top             =   9720
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   15000
      Picture         =   "NEWNEWaTENDIMENTO.frx":330A
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   25
      Top             =   7320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Fechar todas as Seções"
      Height          =   495
      Left            =   240
      TabIndex        =   23
      Top             =   8400
      Width           =   3375
   End
   Begin VB.Timer Timer1 
      Left            =   18360
      Top             =   1680
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   22
      Top             =   8880
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&F1                   Escolher"
      Height          =   735
      Left            =   14760
      Picture         =   "NEWNEWaTENDIMENTO.frx":3D0C
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   2760
      Width           =   1815
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0FFC0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   14760
      TabIndex        =   19
      Top             =   1800
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14760
      TabIndex        =   17
      Top             =   1080
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14520
      TabIndex        =   15
      Top             =   360
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      MaxLength       =   4
      TabIndex        =   1
      Top             =   360
      Width           =   2895
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9480
      MaxLength       =   4
      TabIndex        =   2
      Top             =   360
      Width           =   3135
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      MaxLength       =   65
      TabIndex        =   0
      Top             =   1080
      Width           =   6015
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00C0C0FF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4200
      TabIndex        =   5
      Top             =   1800
      Visible         =   0   'False
      Width           =   9615
   End
   Begin VB.TextBox Text9 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12240
      TabIndex        =   4
      Top             =   1080
      Width           =   1935
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   6735
      Left            =   4200
      Negotiate       =   -1  'True
      TabIndex        =   3
      Top             =   2760
      Visible         =   0   'False
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   11880
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   19
      WrapCellPointer =   -1  'True
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
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Selecione o produto "
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
   Begin ComctlLib.TreeView TreeView1 
      Height          =   8535
      Left            =   240
      TabIndex        =   6
      Top             =   840
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   15055
      _Version        =   327682
      Style           =   7
      Appearance      =   1
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
   Begin VB.TextBox Text16 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   16560
      TabIndex        =   44
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "Entrega"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   255
      Left            =   5400
      TabIndex        =   73
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label21 
      Caption         =   "REVALIDE"
      Height          =   255
      Left            =   19920
      TabIndex        =   70
      Top             =   3360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Edição"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   69
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label19 
      Caption         =   "Edição"
      Height          =   255
      Left            =   20160
      TabIndex        =   68
      Top             =   2640
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label18 
      Caption         =   "entregaouBusca"
      Height          =   255
      Left            =   18240
      TabIndex        =   65
      Top             =   2520
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label17 
      Caption         =   "&F5"
      Height          =   255
      Left            =   13440
      TabIndex        =   63
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label16 
      Caption         =   "ediçao"
      Enabled         =   0   'False
      Height          =   255
      Left            =   17040
      TabIndex        =   61
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label LabelFRETE 
      Caption         =   "Frete aqui! pro"
      Height          =   375
      Left            =   17760
      TabIndex        =   59
      Top             =   6360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label15"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   52
      Top             =   480
      Visible         =   0   'False
      Width           =   3375
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label14 
      Caption         =   "Nº Pedido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   51
      Top             =   120
      Visible         =   0   'False
      Width           =   2895
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label13 
      Caption         =   "Qtd"
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
      Left            =   12840
      TabIndex        =   42
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label12 
      Caption         =   "Total "
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
      Left            =   14760
      TabIndex        =   38
      Top             =   6720
      Width           =   1575
   End
   Begin VB.Label Label12 
      Caption         =   "Frete"
      Height          =   375
      Index           =   0
      Left            =   14760
      TabIndex        =   35
      Top             =   5640
      Width           =   1575
   End
   Begin VB.Label Label11 
      Height          =   375
      Left            =   480
      TabIndex        =   24
      Top             =   480
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label Label10 
      Caption         =   "S_TL"
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
      Left            =   14160
      TabIndex        =   20
      Top             =   2040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label9 
      Caption         =   "½"
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
      Left            =   14400
      TabIndex        =   18
      Top             =   1200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label8 
      Caption         =   "Tam"
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
      Left            =   14520
      TabIndex        =   16
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "Codigo do sistema"
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
      Left            =   8400
      TabIndex        =   14
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "R$"
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
      Left            =   11760
      TabIndex        =   13
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label7 
      Caption         =   "Meu Codigo"
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
      Left            =   4200
      TabIndex        =   12
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Produto descrição"
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
      Left            =   4200
      TabIndex        =   11
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Codigo do sistema"
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
      Left            =   5520
      TabIndex        =   10
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "R$"
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
      Left            =   7440
      TabIndex        =   9
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "Codigo do sistema"
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
      TabIndex        =   8
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "R$"
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
      Left            =   7440
      TabIndex        =   7
      Top             =   1080
      Width           =   375
   End
End
Attribute VB_Name = "Form401"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim gloValor1 As Double
Dim gloValor2 As Double
Dim gloPassos As Integer
Dim gloPrimerioValorParaPrimeirMedida As String
Dim contadorFunCall As Boolean
Dim cardapioLateral As Boolean
Dim Pesquisacardapiofun As Integer
Dim completarFinalizar As Double
Dim numerodopedido As Integer
Dim ultimomultiplicador As Double
Dim novomultiplicador As Double
Dim valorOriginal1 As Double
Dim valorOriginal2 As Double
Dim valorOriginal3 As Double
Dim valorOriginal4 As Double
Dim qtdsup As String
Dim valorFRete As Double
Dim Obsevacao As String
Dim deedlock As Integer
Dim BlocRevalide As Boolean
Dim i As Integer
Dim valorAcrec As Double
Dim tamanHoEscolidoprimeiro As String



Private Sub Command1_Click()
If Text22.Text <> "" Then
If Text31.Text = 0 Then
BlocRevalide = True
Else
BlocRevalide = False
End If
'form500.show
Form500.ProgressBar1.Value = 50

FinalizacaoPontoFinal
Else
MsgBox "Confirme seu pedido novamente", , "Reconfirme o pedido!"
End If
End Sub

Private Sub Command10_Click()
Dim valorcomAcress As Double

'ultimomultiplicador = Text15.Text
Form500.ProgressBar1.Value = 50

'valorOriginal2 = Text4.Text
'valorOriginal1 = Text9.Text
qtdsup = InputBox("Digite a quantidade ideal para o produto", "Quantidade", Text15.Text)

If qtdsup <> "0" Then
If qtdsup <> "" And Text18.Text <> "Text18" Then
Text17 = CDbl(Text18 * qtdsup)
Text15.Text = qtdsup
'valorcomAcress = Replace(Text17, ".", ",")
'Text17.Text = valorcomAcress * qtdsup
'Call Text15_KeyPress(13)
Else
Text15.Text = 1
'Call Text15_KeyPress(13)

'Call Text15_KeyPress(13)


End If
End If

End Sub

Private Sub Command11_Click()
'deedlock = Rnd(50) * 5000
''Sleep (deedlock)

Dim varDbfrete As Double
Dim NomeDaloja As String
Dim NUMEROnoVOrEGISTOobTITO As Integer
'ConServer
NomeDaloja = Form2.DataCombo1.Text

Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs.ActiveConnection = con

rs.CursorLocation = adUseClient

'tratar o valor do frete
If Text12.Text <> "" Then
varDbfrete = CDbl(Text12.Text)
End If
'verificar numero disponivel pedido cancelados
sql = "INSERT INTO `at_contadorDePedidos` (`id`, `intPedido`, `contador`, `situacaoImpressao`) VALUES (NULL, NULL, '0', '0')"
rs.Open sql
'rs.Close
'dead look
''Sleep deedlock
'caso essa funcao esteja procurando por um numerdor
'crie um novo numerador disponivel
sql = "SELECT * FROM `at_contadorDePedidos` WHERE `contador` = 0   ORDER BY `contador` ASC LIMIT 1"
NUMEROnoVOrEGISTOobTITO = 1
Text26 = NUMEROnoVOrEGISTOobTITO
rs.Open sql
If rs.BOF = False And Text29.Text = "0" Then
numerodopedido = rs.Fields("id").Value
  Label14.Visible = True
  Label15.Visible = True
  'registrar contador como usado
  rs.Close
  sql = "UPDATE `robofi61_order_taker`.`at_contadorDePedidos` SET `intPedido` = '" & Form1.StatusBar1.Panels(2).Text & "', `contador` = '50' WHERE (`id` = '" & numerodopedido & " ')"

rs.Open sql
'rs.Close
sql = "SELECT * FROM `at_contadorDePedidos` WHERE `contador` = 0 ORDER BY `at_contadorDePedidos`.`contador` DESC"

rs.Open sql

  Label15.Caption = numerodopedido
End If
   ' rs.Close
    If Text29.Text = 1 Then
   numerodopedido = Label15.Caption
    End If




' inserir numero do pedido

If numerodopedido = 0 Then
Text23.Text = "1"
'sql = "INSERT INTO `at_contadorDePedidos` (`id`, `intPedido`, `contador`) VALUES (NULL, '" & NomeDaloja & "', '1');"
sql = "INSERT INTO `at_contadorDePedidos` (`id`, `intPedido`, `contador`, `situacaoImpressao`) VALUES (NULL, '" & NomeDaloja & "', '0', '100)"
NUMEROnoVOrEGISTOobTITO = 1
Text26 = NUMEROnoVOrEGISTOobTITO
rs.Open sql

  'apanhar numero do  PEDIDO OFICIAL
        
         sql = "SELECT * FROM `at_contadorDePedidos`ORDER BY `id` DESC LIMIT 1"
         rs.Open sql
        
  numerodopedido = rs.Fields("id").Value
  Label14.Visible = True
  Label15.Visible = True
  Label15.Caption = numerodopedido
    rs.Close
    
    
    
    
End If
'verificar se o frete do pedido foi retirado
'sql = "INSERT INTO `at_frete` (`id`, `frete`, `fk_numPedido`) VALUES (NULL, '" & Text20.Text & "', '" & numerodopedido & "')"
'  sql = "SELECT * FROM `at_frete` WHERE `fk_numPedido` = '" & numerodopedido & "' ORDER BY `fk_numPedido` DESC"
rs.Close
 sql = "SELECT * FROM `at_frete` WHERE `frete` = 0 AND `fk_numPedido` = '" & numerodopedido & "' ORDER BY `fk_numPedido` DESC"
  rs.Open sql
  If rs.BOF = False Then
  rs.Close
  If Text29.Text <> 1 Then
  sql = "UPDATE `at_frete` SET `frete` = '" & Text20.Text & "' WHERE `at_frete`.`fk_numPedido` = '" & numerodopedido & "'"
  rs.Open sql
  End If
'  rs.Close
  Else
  rs.Close
  End If

'verificar se e entrega ou buscar
 If Text12.Text <> "R$ 0,00" Then
 Text30.Text = 1
 'entregar
 Else
  Text30.Text = 2
 
 'buscar
 End If



'inserindo frete
' If Text12.Text <> "R$ 0,00" And
'  rs.Close
  If Text29.Text <> 1 Then
  'sql = "INSERT INTO `at_frete` (`id`, `frete`, `fk_numPedido`) VALUES (NULL, '" & Text20.Text & "', '" & numerodopedido & "')"
  'rs.Open sql

  'Text12.Text = Format(varDbfrete, "currency")

  
  
End If
Dim operador As String
'trarar erro
On Error GoTo error

operador = Form1.StatusBar1.Panels(2).Text

 
'usage
''Sleep 3000
If Text15.Text = 1 Then
'Call Command10_Click
End If
Set rs.ActiveConnection = con
' sql = "INSERT INTO `at_itens` ( `Quantidade`, `descrição`, `valor`, `atendente_fk`, `dataTimer`, `fk_pedido`, `fk_cliente`,`observacao`) VALUES ( '" & Text15.Text & "',   '" & Text22.Text & "',  '" & Text17.Text & "', '" & operador & "', '" & Now & "', '" & numerodopedido & "', '" & Form2.Text1.Text & "', '" & Obsevacao & "')"
 ' rs.Open sql
  'Sleep 3000
  Set rs = Nothing
   'Text15.Text = 1
    Obsevacao = ""
Exit Sub

error:
 Set rs = Nothing
'ConServer




Set rs.ActiveConnection = con

rs.CursorLocation = adUseClient
'Sleep 3000


'rs.Close
'sql = "INSERT INTO `at_itens` ( `Quantidade`, `descrição`, `valor`, `atendente_fk`, `dataTimer`, `fk_pedido`, `fk_cliente`,`observacao`) VALUES ( '" & Text15.Text & "',   '" & Text22.Text & "',  '" & Text17.Text & "', '" & operador & "', '" & Now & "', '" & numerodopedido & "', '" & Form2.Text1.Text & "', '" & Obsevacao & "')"
 ' rs.Open sql
  'Text15.Text = 1
Obsevacao = ""
 Set rs = Nothing

Exit Sub
  
  Set rs = Nothing


End Sub

Private Sub Command12_Click()
If Text32.Text = 0 Then
Text32.Text = 1
Else
Text32.Text = 0
End If
End Sub

Private Sub Command13_Click()
cobreOvalormaisAlto
End Sub

Private Sub Command2_Click()
criaritens
Text8.Text = ""
Pesquisacardapiofun = 0
End Sub

Private Sub Command3_Click()
'deedlock = Rnd(50) * 5000
''Sleep (deedlock)

Dim varDbfrete As Double
Dim NomeDaloja As String
Dim NUMEROnoVOrEGISTOobTITO As Integer
'ConServer
NomeDaloja = Form2.DataCombo1.Text

Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs.ActiveConnection = con

rs.CursorLocation = adUseClient

'tratar o valor do frete
If Text12.Text <> "" And Text12.Text <> "Frete aqui! pro" Then
varDbfrete = CDbl(Text12.Text)
End If
'verificar numero disponivel pedido cancelados
sql = "INSERT INTO `at_contadorDePedidos` (`id`, `intPedido`, `contador`, `situacaoImpressao`) VALUES (NULL, NULL, '0', '0')"
rs.Open sql
'rs.Close
'dead look
''Sleep deedlock
'caso essa funcao esteja procurando por um numerdor
'crie um novo numerador disponivel
If BlocRevalide = True And Text32.Text = 0 Then
            sql = "SELECT * FROM `at_contadorDePedidos` WHERE `contador` = 0   ORDER BY `contador` ASC LIMIT 1"
            NUMEROnoVOrEGISTOobTITO = 1
            Text26 = NUMEROnoVOrEGISTOobTITO
            rs.Open sql
            If rs.BOF = False And Text29.Text = "0" Then
            numerodopedido = rs.Fields("id").Value
              Label14.Visible = True
              Label15.Visible = True
              'registrar contador como usado
              rs.Close
              sql = "UPDATE `robofi61_order_taker`.`at_contadorDePedidos` SET `intPedido` = '" & Form1.StatusBar1.Panels(2).Text & "', `contador` = '50' WHERE (`id` = '" & numerodopedido & " ')"
            
            rs.Open sql
            'rs.Close
            sql = "SELECT * FROM `at_contadorDePedidos` WHERE `contador` = 0 ORDER BY `at_contadorDePedidos`.`contador` DESC"
            
            rs.Open sql
            
              Label15.Caption = numerodopedido
              rs.Close
   End If
Else
 numerodopedido = Label15.Caption

End If
   
   ' rs.Close
    If Text29.Text = 1 Then
   numerodopedido = Label15.Caption
    End If




' inserir numero do pedido

If numerodopedido = 0 Then
Text23.Text = "1"
'sql = "INSERT INTO `at_contadorDePedidos` (`id`, `intPedido`, `contador`) VALUES (NULL, '" & NomeDaloja & "', '1');"
sql = "INSERT INTO `at_contadorDePedidos` (`id`, `intPedido`, `contador`, `situacaoImpressao`) VALUES (NULL, '" & NomeDaloja & "', '0', '100)"
NUMEROnoVOrEGISTOobTITO = 1
Text26 = NUMEROnoVOrEGISTOobTITO
rs.Open sql

  'apanhar numero do  PEDIDO OFICIAL
        
         sql = "SELECT * FROM `at_contadorDePedidos`ORDER BY `id` DESC LIMIT 1"
         rs.Open sql
        
  numerodopedido = rs.Fields("id").Value
  Label14.Visible = True
  Label15.Visible = True
  Label15.Caption = numerodopedido
    rs.Close
    
    
    
    
End If
'verificar se o frete do pedido foi retirado
'sql = "INSERT INTO `at_frete` (`id`, `frete`, `fk_numPedido`) VALUES (NULL, '" & Text20.Text & "', '" & numerodopedido & "')"
'  sql = "SELECT * FROM `at_frete` WHERE `fk_numPedido` = '" & numerodopedido & "' ORDER BY `fk_numPedido` DESC"
'rs.Close
 sql = "SELECT * FROM `at_frete` WHERE `frete` = 0 AND `fk_numPedido` = '" & numerodopedido & "' ORDER BY `fk_numPedido` DESC"
  rs.Open sql
  If rs.BOF = False Then
  rs.Close
  If Text29.Text <> 1 Then
  sql = "UPDATE `at_frete` SET `frete` = '" & Text20.Text & "' WHERE `at_frete`.`fk_numPedido` = '" & numerodopedido & "'"
  rs.Open sql
  End If
'  rs.Close
  Else
  rs.Close
  End If

'verificar se e entrega ou buscar
 If Text12.Text <> "R$ 0,00" Then
 Text30.Text = 1
 'entregar
 Else
  Text30.Text = 2
 
 'buscar
 End If



'inserindo frete
' If Text12.Text <> "R$ 0,00" And
'  rs.Close
  If Text29.Text <> 1 Then
  sql = "INSERT INTO `at_frete` (`id`, `frete`, `fk_numPedido`) VALUES (NULL, '" & Text20.Text & "', '" & numerodopedido & "')"
  rs.Open sql

  Text12.Text = Format(varDbfrete, "currency")

  
  
End If
Dim operador As String
'trarar erro
On Error GoTo error

operador = Form1.StatusBar1.Panels(2).Text

 
'usage
'Sleep 3000
If Text15.Text = 1 And Text17.Text <> "0" Then
Call Command10_Click
End If
Set rs.ActiveConnection = con
 sql = "INSERT INTO `at_itens` ( `Quantidade`, `descrição`, `valor`, `atendente_fk`, `dataTimer`, `fk_pedido`, `fk_cliente`,`observacao`) VALUES ( '" & Text15.Text & "',   '" & Text22.Text & "',  '" & Text17.Text & "', '" & operador & "', '" & Now & "', '" & numerodopedido & "', '" & Form2.Text1.Text & "', '" & Obsevacao & "')"
  rs.Open sql
              Text16.Visible = False
            Text16 = 0
  qtdsup = 1
  'catedaral
 ' 'Sleep 3000
 
  ' Set rs = Nothing
   
 '
   'Text15.Text = 1
    Obsevacao = ""
Exit Sub

error:
 Set rs = Nothing
ConServer




Set rs.ActiveConnection = con

rs.CursorLocation = adUseClient
'Sleep 3000
If Text17.Text <> "0" Then

'rs.Close
sql = "INSERT INTO `at_itens` ( `Quantidade`, `descrição`, `valor`, `atendente_fk`, `dataTimer`, `fk_pedido`, `fk_cliente`,`observacao`) VALUES ( '" & Text15.Text & "',   '" & Text22.Text & "',  '" & Text17.Text & "', '" & operador & "', '" & Now & "', '" & numerodopedido & "', '" & Form2.Text1.Text & "', '" & Obsevacao & "')"
  rs.Open sql
              Text16.Visible = False
            Text16 = 0
  qtdsup = 1
  End If
Obsevacao = ""
 Set rs = Nothing

Exit Sub
  
  Set rs = Nothing

End Sub

Private Sub Command4_Click()
Animation1.Close
End Sub



Private Sub Command5_Click()
Form402.Visible = True
Text16.Visible = True
If Text7.Visible = True Then
Form402.Caption = " Acréssimo para  " & Text7.Text
Else
Form402.Caption = " Acréssimo para  " & Text3.Text
End If

End Sub

Private Sub Command6_Click()
Form403.Visible = True
If Text7.Visible = True Then
Form403.Caption = " Observações  para  " & Text7.Text
Else
Form403.Caption = " Observações para  " & Text3.Text
End If
contabilize
End Sub

Private Sub Command7_Click()
Dim resp As Integer
 If Text2.Text = "" Then
    'CommonDialog1.CancelError = True
    'trarar erro
    On Error GoTo error
    
    valorOriginal1 = 0
    valorOriginal2 = 0
    novomultiplicador = 0
    
    
    Form3.Hide
    Dialog.Hide
    'form500.show
        'ConServer
    
    
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    Set rs.ActiveConnection = con
    
    Dim valorDosItens As Double
    
    Form500.ProgressBar1.Value = 30
    Form404.Visible = True
    Form404.Text2.Text = Text28.Text & " " & Form404.Text2.Text
    
    Form404.Label20.Caption = Format(Text25.Text, ">")
    'data hora
    Form404.Label19 = Now
    'operador
    'Form404.Label20 = ""
    'observaçoes
    If Text13.Text <> "" Then
    Form404.Text2.Text = Text13.Text + Form404.Text2.Text
    Form500.ProgressBar1.Value = 40
    End If
    Form404.Caption = "Pedido Nº" & Label15
    Form404.Label18.Caption = numerodopedido
    
    Form404.Adodc1.RecordSource = ""
    
    Form404.Adodc1.CommandType = adCmdText
    
    Form404.Adodc1.RecordSource = "SELECT `Quantidade`,`descrição`,`valor` ,`id`FROM `at_itens` WHERE `fk_pedido` = '" & numerodopedido & "'"
    
    Form404.Adodc1.Refresh
    ''DataGrid1.Columns(0).Caption = NpTitulo
    Form404.DataGrid1.Columns(1).Width = 3800
    'buscar o frete
    Form500.ProgressBar1.Value = 50
    Form404.Adodc2.RecordSource = ""
    
    Form404.Adodc2.CommandType = adCmdText
    
    Form404.Adodc2.RecordSource = "SELECT `frete` FROM `at_frete` WHERE `fk_numPedido` ='" & numerodopedido & "'ORDER BY `at_frete`.`id` DESC "
    
    Form404.Adodc2.Refresh
    Form500.ProgressBar1.Value = 60
    If Form404.Adodc2.Recordset.BOF = False Then
    
    
    
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
    Form401.Hide
    Form404.Visible = True
    Form404.SetFocus



Exit Sub

error:

valorOriginal1 = 0
valorOriginal2 = 0
novomultiplicador = 0


Form3.Hide
Dialog.Hide
'form500.show
    ConServer



Set rs = New ADODB.Recordset

Set rs.ActiveConnection = con


Form500.ProgressBar1.Value = 30
Form404.Show
Form404.Text2.Text = Text28.Text & " " & Form404.Text2.Text

Form404.Label20.Caption = Format(Text25.Text, ">")
'data hora
Form404.Label19 = Now
'operador
'Form404.Label20 = ""
'observaçoes
If Text13.Text <> "" Then
Form404.Text2.Text = Text13.Text + Form404.Text2.Text
Form500.ProgressBar1.Value = 40
End If
Form404.Caption = "Pedido Nº" & Label15
Form404.Label18.Caption = numerodopedido

Form404.Adodc1.RecordSource = ""

Form404.Adodc1.CommandType = adCmdText

Form404.Adodc1.RecordSource = "SELECT `Quantidade`,`descrição`,`valor` ,`id`FROM `at_itens` WHERE `fk_pedido` = '" & numerodopedido & "'"

Form404.Adodc1.Refresh
''DataGrid1.Columns(0).Caption = NpTitulo
Form404.DataGrid1.Columns(1).Width = 3800
'buscar o frete
Form500.ProgressBar1.Value = 50
Form404.Adodc2.RecordSource = ""

Form404.Adodc2.CommandType = adCmdText

Form404.Adodc2.RecordSource = "SELECT `frete` FROM `at_frete` WHERE `fk_numPedido` ='" & numerodopedido & "'ORDER BY `at_frete`.`id` DESC "

Form404.Adodc2.Refresh
Form500.ProgressBar1.Value = 60
If Form404.Adodc2.Recordset.BOF = False Then



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
Form401.Hide
Form404.Visible = True
Form404.SetFocus

Exit Sub
'Call Form404.Command7_Click


Else


resp = MsgBox("Seu pedido presente na tela de " & Text3.Text & " ainda não foi incluido deseja proceguir mesmo assim", vbYesNo + vbCritical, "Pedido não Finalizado")
If resp = 6 Then
'aviso para limpar caso queria um novo agregamento
finalizarLimparProceguir

Else
'manter o quadro
End If
End If

End Sub

Private Sub DataGrid1_Click()
'organizadoDePassos (gloPassos)
Text3.Text = ""
Text5.Text = DataGrid1.Columns(3).Value
 If Not IsEmpty(DataGrid1.Columns(1).Value) Then
Text1.Text = DataGrid1.Columns(1).Value
End If
Text2.Text = DataGrid1.Columns(0).Value

Text3.Text = DataGrid1.Columns(2).Value

Text9.Text = DataGrid1.Columns(5).Value





Call Text3_KeyPress(13)
End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)

  If KeyCode = vbKeyF1 Then
    Call Command1_Click
    'Text15.Text = 1
    
  End If
    If KeyCode = vbKeyF2 Then
    If Command5.Enabled = True Then
   Call Command5_Click
   Else
   MsgBox "Antes escolha o produto", , "Produto?"
   End If
  End If
    If KeyCode = vbKeyF3 Then
   If Command6.Enabled = True Then
   Call Command6_Click
   Else
   MsgBox "Antes escolha o produto", , "Produto?"
   End If
   End If
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

Call DataGrid1_Click

End If
End Sub

Private Sub DataGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    If Text7.Visible = True Then
        Text7.Visible = False
       
        ElseIf Text7.Visible = False Then
         Text7.Visible = True
         completarFinalizar = 0
         Text10.Text = completarFinalizar
        Text7.Text = ""
        End If
refuncall

End If

 

End Sub



Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 47 Or KeyAscii = 92 Then
KeyAscii = 0
'ultimomultiplicador = 0

        If Text7.Visible = True Then
        Text7.Visible = False
       
        ElseIf Text7.Visible = False Then
         Text7.Visible = True
         completarFinalizar = 0
         Text10.Text = completarFinalizar
        Text7.Text = ""
        Text3.SetFocus
        SendKeys "{ENTER}"
        End If
refuncall

End If

 If KeyAscii = vbKeyEscape Then
 'form500.show
 Text4.Text = 0
   finalizar
    End If

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF1 Then
    Call Command1_Click
  End If
    If KeyCode = vbKeyF2 Then
    If Command5.Enabled = True Then
   Call Command5_Click
   Else
   MsgBox "Antes escolha o produto", , "Produto?"
   End If
  End If
    If KeyCode = vbKeyF3 Then
     If Command6.Enabled = True Then
   Call Command6_Click
   Else
   MsgBox "Antes escolha o produto", , "Produto?"
   End If
  End If
    If KeyCode = vbKeyF4 And Command7.Enabled = True Then
    
    Call Command7_Click
    
  End If
      If KeyCode = vbKeyF5 Then
    
    Call Command10_Click
    
  End If

End Sub

Private Sub Form_Load()
Form404.Show
'Form23.Visible = True
Form404.Visible = False
'deedlock = Rnd(50) * 5000
'Form2.Hide
 'Form1.StatusBar1.Panels(2).Text = "testing"
 

Text25.Text = Form1.StatusBar1.Panels(2).Text


Form2.Command2.Enabled = True
completarFinalizar = 0
ultimomultiplicador = 1
novomultiplicador = 0

Text10.Text = completarFinalizar
 contadorFunCall = False
gloPassos = 1
criaritens

Form2.Animation1.Close
'Form2.Hide
Form500.Hide
Dialog.Hide
End Sub
Public Sub criaritens()
Dim totalDecardapiosListados As Integer
Dim coutCardapioTitulo As Integer
Dim procurarSub As String

ConServer


Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs.ActiveConnection = con

rs.CursorLocation = adUseClient
  sql = "SELECT COUNT(DISTINCT `Titulo`) FROM `Cardapio` WHERE 1"
  rs.Open sql
  coutCardapioTitulo = rs.Fields("COUNT(DISTINCT `Titulo`)").Value

Dim nodx As Node

'limpa qualquer nó criado
TreeView1.Nodes.Clear

'Set TreeView1.ImageList = ImageList1

'novabuscapelos titulos
  rs.Close
  sql = "SELECT DISTINCT`Titulo` FROM `Cardapio` WHERE 1"
  rs.Open sql
 'coutCardapioTitulo = rs.Fields("COUNT(DISTINCT `Titulo`)").Value

For i = coutCardapioTitulo To 1 Step -1
 

'Adicionar titulo
 Dim sKey As String
    Dim oNodex As Node
    
''Adicione alguns itens de nível raiz
''
'para i = 1 a 5
'TreeView1.Nodes.Add , , "ROOT" & i, "Item Raiz" & i
TreeView1.Nodes.Add , , "ROOT" & i, rs.Fields("Titulo").Value
totalDecardapiosListados = totalDecardapiosListados + 1
rs.MoveNext


Next

'próximo
'"
'Agora adicionar alguns filhos
'"
'para i = 1 a 5
'com TreeView1.Nodes
'.Adicionar "ROOT1", tvwChild, "ROOT1CHILD" & i, "Item infantil" & i
'.Adicionar "ROOT2", tvwChild, "ROOT2CHILD" & i, "Item infantil" & i
'.Adicionar "ROOT3", tvwChild, "ROOT3CHILD" & i, "Item para crianças" & i
'.Adicione "ROOT4", tvwChild, "ROOT4CHILD" & i, "Item infantil" & i
'.Adicione "ROOT5", tvwChild, "ROOT5CHILD" & i, "Item para crianças" & i
'End With
'Next
''
' definir  o quantos subs tem para serem creidos
Dim t As Integer
Dim redusCardapios As Integer
redusCardapios = totalDecardapiosListados
For t = 1 To totalDecardapiosListados
                 
                procurarSub = TreeView1.Nodes.Item(t)
                rs.Close
                  sql = "SELECT DISTINCT `codTitulo` FROM `Cardapio` WHERE `Titulo` LIKE '" & procurarSub & "' "
                  rs.Open sql
                  Dim retornaCodigodoiten As Integer
                  
                  retornaCodigodoiten = rs.Fields("codTitulo").Value
                 rs.Close
                  sql = "SELECT COUNT(DISTINCT `tipo`) FROM `Cardapio` WHERE `codTitulo`='" & retornaCodigodoiten & "'"
                  rs.Open sql
                  coutCardapioTitulo = rs.Fields("COUNT(DISTINCT `tipo`)").Value
                
                
                rs.Close
                 sql = "SELECT DISTINCT `tipo` FROM `Cardapio` WHERE `codTitulo`= '" & retornaCodigodoiten & "'"
                rs.Open sql
                For i = 1 To coutCardapioTitulo
                With TreeView1.Nodes
                
                If (rs.Fields("tipo").Value <> "") Then
                .Add "ROOT" & redusCardapios, tvwChild, "ROOT" & redusCardapios & "CHILD" & i, rs.Fields("tipo").Value
                 End If
                End With
                rs.MoveNext
                Next
redusCardapios = redusCardapios - 1
Next

'' Agora adicione alguns Grand- Children
''
'para i = 1 a 5
'com TreeView1.Nodes
'.Add "ROOT1CHILD2", tvwChild, "Grand criança" & i
'.Add "ROOT2CHILD2", tvwChild, "Grand criança" & i
'.Add "ROOT3CHILD2", tvwChild, , "Grand filho" & i
'.Adicionar "ROOT4CHILD2", tvwChild, "Grand filho" & i
'.Adicione "ROOT5CHILD2", tvwChild, "Grand Child" e termino
'com o
'próximo
'For i = 1 To 5
'With TreeView1.Nodes
'.Add "ROOT1CHILD2", tvwChild, , "Grand Child " & i
'.Add "ROOT2CHILD2", tvwChild, , "Grand Child " & i
'.Add "ROOT3CHILD2", tvwChild, , "Grand Child " & i
'.Add "ROOT4CHILD2", tvwChild, , "Grand Child " & i
'.Add "ROOT5CHILD2", tvwChild, , "Grand Child " & i
'End With
'Next














 Set rs = Nothing

     
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload Form23
Dim resp As Integer
If Text26.Text = 0 Then
resp = MsgBox(" Você não fechou o pedido por completo ,desta forma o pedido será cancelado ! deseja mesmo fazer isso ?", vbYesNo, "Cancelamento do pedido " & Label15.Caption)
If resp = 6 Then
CancelarPedido
Form401.Text23.Text = 0
Form401.Text26.Text = 0
Else
Form401.Text23.Text = 0
Form401.Text26.Text = 0
End If
End If
Form2.Show
Form2.Command2.Enabled = True
Form500.Hide
Form23.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form23.Visible = False
End Sub

Private Sub Label15_Change()
If Label15.Visible = True Then
'verificar se o numero nao foi escolhido por outra maquina
If BlocRevalide = True Then
Revalide
Else
Label14.Visible = True
End If
Text26.Text = 0
Command7.Enabled = True
Else
Command7.Enabled = False
End If
End Sub

Private Sub Text1_Change()
Text15.Text = 1
If Text1.Text = "" Then
Text2.Text = ""
Text3.Text = ""

cobreOvalormaisAlto
REPARSEvalores
LimparUltimaescolha
If completarFinalizar = 0 Then
Text5.Text = ""
Text9.Text = ""
End If


End If
End Sub

Private Sub Text1_DblClick()
Text1.Text = ""
End Sub

Private Sub Text1_GotFocus()
'MsgBox Pesquisacardapiofun
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Then
Text3.Text = ""
Text3.SetFocus
End If
If KeyCode = 39 Then
Text2.Text = ""
Text2.SetFocus
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = (SoNumeros(KeyAscii))
' AO PRECIONAR ENTER NA CAIXA

If KeyAscii = 13 Then
KeyAscii = 0

limparbasedecontas
Animation1.Open Text11.Text

Animation1.AutoPlay = True
    funCardapioPorMeuCodigo
  Call Command13_Click
End If

End Sub
Public Sub organizadoDePassos(passo As Integer)





  'Sleep 1000
Select Case passo
Case 1
'definirmetades
If DataGrid1.Visible = True Then
    
    If Text3.Text = (Trim(DataGrid1.Columns(0).Value)) Then
    If FunCardapioPordescricao(Trim(Text3.Text)) = True Then
              
               If definaAsmedidaspossiveis(Trim(DataGrid1.Columns(0).Value)) = True Then
               gloPassos = 2
               End If
            End If
    Else
    Text3.Text = (Trim(DataGrid1.Columns(0).Value))
    End If
End If
    
    'contadorDePassos = contadorDePassos + 1
Case 2

   saberOcodigo

'saberOPreco
Case 3
If DataGrid1.Visible = True Then
    If Text3.Text = (DataGrid1.Columns(0).Value) Then
    definaAsmedidaspossiveis (DataGrid1.Columns(0).Value)
    'definaAsmedidaspossiveis(Mcodigo As String)
    'saberOcodigo
    Else
    Text3.Text = (DataGrid1.Columns(0).Value)
    Text7.Visible = True
    
   
    
    If completarFinalizar = 0.5 And Text7.Text <> "" Then
Text7.Text = Text7.Text & "   1/2 " & Text3.Text
Else
 Text7.Text = "1/2  " & Text3.Text
End If
    
    'saberOcodigo
    End If
End If





'requerer2 entradas
'valores pela metade
'armazenar valores



 
 
Case 4
 'verificar seleçao
If Text5.Text = (DataGrid1.Columns(0).Value) Then
    'criar novo evento
    MsgBox "incluir esse produto na venda"
    
   Else
   Text5.Text = (DataGrid1.Columns(0).Value)
  
  
   End If
'ja e possivel saber o codigo da medida escolhida
'já e possivel saber o preço

'saberOPreco


'finalizacaodemetates



Case 5
'prepara reinicio

   'primeiraMetadeDaLarnja = Text7.Text
   DataGrid1.Visible = False
   Text3.Text = ""
   Text2.Text = ""
   Text1.Text = ""
   Text7.Visible = True
   Text3.SetFocus
'contadorDePassos = 6
Text7.Visible = True
'organizadoDePassos (contadorDePassos)

 
Case 6
If DataGrid1.Visible = True Then
Text4.Visible = True
Label9.Visible = True
 'reiniciar
     If Text3.Text = (DataGrid1.Columns(0).Value) Then
  
  'saberOPrecoDaOutraMetade
    Else
    Text3.Text = (DataGrid1.Columns(0).Value)
    Text7.Visible = True
    'Text7.Text = primeiraMetadeDaLarnja & " 1/2  " & Text3.Text
    
    End If

End If
Case 7
 'MOVIMENTO POR MEU CODIGO
 
 
'If ValordaPrimeiraMetade <> 0 And Text7.Visible = False Then
'finalize
'Else
        'If ValordaPrimeiraMetade = 0 Then
         Text7.Visible = True
             If completarFinalizar = 0.5 And Text7.Text <> "" Then
Text7.Text = Text7.Text & "   1/2 " & Text3.Text
Else
 Text7.Text = "1/2  " & Text3.Text
End If
         Text9.Text = Text9 / 2
         'ValordaPrimeiraMetade = Text9.Text
         'Text9.Text = Format(ValordaPrimeiraMetade, "currency")
         Text3.SetFocus
         'Else
    '     contadorDePassos = 8
         Text7.Visible = True
   '      organizadoDePassos (contadorDePasso)
         Text3.SetFocus
         'End If
         
 'End If
 
 
 
 Case 8
 ' preparando novo entrada segundo iten
 If DataGrid1.Visible = True Then
 Text3.Text = DataGrid1.Columns(0).Value
 Else
 Text3.Text = ""
 
 End If
   'primeiraMetadeDaLarnja = Text7.Text

   
   Text2.Text = ""
   Text1.Text = ""
   Text7.Visible = True
   Text3.SetFocus
 'tamanho definido
 
 
 
 
 
 
 Case 9
 ' conclusao
    If completarFinalizar = 0.5 And Text7.Text <> "" Then
Text7.Text = Text7.Text & "   1/2 " & Text3.Text
Else
 Text7.Text = "1/2  " & Text3.Text
End If
 DataGrid1.Visible = False
 MsgBox "fim do pedido metades"
 Text3.SetFocus
 
 
 
 Case 10
 Case 11
 Case 12
 Case 13
 
 
 
End Select


End Sub





' CARDAPIO POR  SISTEMA
Public Sub CardapioCodigosistem(Mcodigo As Integer)
Dim intMcodigo As Integer
Dim valorObtido As Double
intMcodigo = Mcodigo

'ConServer

Dim NpTitulo As String

Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs.ActiveConnection = con
rs.CursorLocation = adUseClient

Select Case Pesquisacardapiofun
            Case 0 'simples pesquisa
            sql = "SELECT `idCardapio`,`codigo`,`Descricao` ,`tipo`,`valor`,`Medida`FROM `Cardapio` WHERE `idCardapio` = '" & intMcodigo & "'ORDER BY `Cardapio`.`Descricao` ASC "
                                                rs.Open sql
                                                 If rs.BOF = False Then
                                                   If Not IsNull(rs.Fields("codigo").Value) Then
                                                  Text1.Text = rs.Fields("codigo").Value
                                                  End If
                                                  Text2.Text = rs.Fields("idCardapio").Value
                                                  Text3.Text = rs.Fields("Descricao").Value
                                                  Text5.Text = rs.Fields("Medida").Value
                                                  funcall (rs.Fields("valor").Value)
                                                  
                                                                                                          
                                                End If
                                                rs.Close
         
  
  sql = "SELECT `Descricao` FROM `Cardapio` WHERE `idCardapio` = '" & intMcodigo & "'ORDER BY `Cardapio`.`Descricao` ASC "
  rs.Open sql
 If rs.BOF = False Then
 'Set DataGrid1.DataSource = rs
 'DataGrid1.Visible = True
 'DataGrid1.SetFocus
 Else
 MsgBox " O codigo : '" & Mcodigo & "'não foi localizado no sistema ", , "Não Existe"
 DataGrid1.Visible = False
 
 End If


'rs.Close
                                                 
'DataGrid1.Columns(0).Caption = NpTitulo
'DataGrid1.Columns(0).Width = 12000

 Set rs = Nothing
              
            Case 1 'pesquisa por titulos
sql = "SELECT `idCardapio`,`codigo`,`Descricao` ,`tipo`,`valor`,`Medida`FROM `Cardapio` WHERE `idCardapio` = '" & intMcodigo & "'AND`Titulo` LIKE'" & Text8.Text & "'ORDER BY `Cardapio`.`Descricao` ASC "
                                                rs.Open sql
                                                 If rs.BOF = False Then
                                                   If Not IsNull(rs.Fields("codigo").Value) Then
                                                  Text1.Text = rs.Fields("codigo").Value
                                                  End If
                                                  Text2.Text = rs.Fields("idCardapio").Value
                                                  Text3.Text = rs.Fields("Descricao").Value
                                                  Text5.Text = rs.Fields("Medida").Value
                                                  funcall (rs.Fields("valor").Value)
                                                  
                                                                                                          
                                                End If
                                                rs.Close
         
  
  sql = "SELECT `Descricao` FROM `Cardapio` WHERE `idCardapio` = '" & intMcodigo & "'AND`Titulo` LIKE'" & Text8.Text & "'ORDER BY `Cardapio`.`Descricao` ASC "
  rs.Open sql
 If rs.BOF = False Then
 'Set DataGrid1.DataSource = rs
 'DataGrid1.Visible = True
 'DataGrid1.SetFocus
 Else
 MsgBox " O codigo : '" & Mcodigo & "'não foi localizado no sistema  no Cardápio " & Text8.Text & "", , "Não Existe"
 DataGrid1.Visible = False
 
 End If


'rs.Close
                                                 
'DataGrid1.Columns(0).Caption = NpTitulo
'DataGrid1.Columns(0).Width = 12000

 Set rs = Nothing
            Case 2 'pesquisar subtitulos
            
sql = "SELECT `idCardapio`,`codigo`,`Descricao` ,`tipo`,`valor`,`Medida`FROM `Cardapio` WHERE `idCardapio` = '" & intMcodigo & "'AND`tipo` LIKE'" & Text8.Text & "'ORDER BY `Cardapio`.`Descricao` ASC "
                                                rs.Open sql
                                                 If rs.BOF = False Then
                                                 If Not IsNull(rs.Fields("codigo").Value) Then
                                                  Text1.Text = rs.Fields("codigo").Value
                                                  End If
                                                  Text2.Text = rs.Fields("idCardapio").Value
                                                  Text3.Text = rs.Fields("Descricao").Value
                                                  Text5.Text = rs.Fields("Medida").Value
                                                  funcall (rs.Fields("valor").Value)
                                                  
                                                                                                          
                                                End If
                                                rs.Close
         
  
  sql = "SELECT `Descricao` FROM `Cardapio` WHERE `idCardapio` = '" & intMcodigo & "'AND`tipo` LIKE'" & Text8.Text & "'ORDER BY `Cardapio`.`Descricao` ASC "
  rs.Open sql
 If rs.BOF = False Then
 'Set DataGrid1.DataSource = rs
 'DataGrid1.Visible = True
 'DataGrid1.SetFocus
 Else
 MsgBox " O codigo : '" & Mcodigo & "'não foi localizado no sistema  , cardápio selecionado indice " & Text8.Text & "", , "Não Existe"
 DataGrid1.Visible = False
 
 End If


'rs.Close
                                                 
'DataGrid1.Columns(0).Caption = NpTitulo
'DataGrid1.Columns(0).Width = 12000

 Set rs = Nothing
        End Select

                                       
End Sub



Private Sub Text12_Change()

If Text12.Text <> "" And Text12.Text <> "R$ 0,00" Then
Text20.Text = Text12.Text
Else
Text12.Visible = False
End If
End Sub

Private Sub Text14_Change()
If Text14.Text <> "" Then
Text17.Text = DBLcurre(Text14.Text)

End If
End Sub

Private Sub Text15_Click()
If Text9.Text = "" Then
MsgBox "Defina primerio o valor do produto!", , "Valor não Definido"
If DataGrid1.Visible = True Then

DataGrid1.SetFocus
Else
Text3.SetFocus
End If
Else
Call Command10_Click
End If
End Sub

Private Sub Text15_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Dim valor1 As Double
Dim valor2 As Double
Dim valori As Double

If novomultiplicador = 0 Then
'Text15.Text = 1
novomultiplicador = 1
End If
If valorAcrec = 0 And valorAcrec = "R$0,00" Then
valor1 = valorOriginal1 / ultimomultiplicador
valor2 = valorOriginal2 / ultimomultiplicador
novomultiplicador = Text15

'multiplica
Text9.Text = Format(valor1 * novomultiplicador, "currency")
Text4.Text = Format(valor2 * novomultiplicador, "currency")
ultimomultiplicador = novomultiplicador
Else
'ache o valor original com acressimo
valor1 = valorOriginal1 / ultimomultiplicador
valor2 = valorOriginal2 / ultimomultiplicador
novomultiplicador = Text15

'multiplica
Text9.Text = Format(valor1 * novomultiplicador, "currency")
Text4.Text = Format(valor2 * novomultiplicador, "currency")
ultimomultiplicador = novomultiplicador


End If




End If
End Sub

Private Sub Text16_Change()
If Text16.Text <> "" And Text16.Text <> "0" And Text16.Text <> "R$ 0,00" Then
Dim comecApartirDe As Integer
Dim sringComoFica As String
'Text19.Text = Text16.Text
Text18.Text = CDbl(Text9) + CDbl(Text16)
reescrevaonomedoProduto
'Text18.Text = Text17
Text16.Visible = True
Text24.Text = Text16.Text

contabilize
Else
Text16.Visible = False
If Text22.Text <> "Text22" Then
comecApartirDe = InStr(3, Text22.Text, " AC ", 0)
sringComoFica = Left(Text22.Text, comecApartirDe)
'comecApartirDe = InStr(3, Text22.Text, " AC RETIRAR ACRÉSSIMO", 0)
'sringComoFica = Left(Text22.Text, comecApartirDe)
Text22.Text = sringComoFica
End If
contabilize
End If
Call Command13_Click
End Sub

Private Sub Text17_Change()
Text17 = DBLcurre(Text17.Text)
End Sub

Private Sub Text18_Change()
Text17 = DBLcurre(Text18.Text)

End Sub

Private Sub Text19_Change()
Text19 = DBLcurre(Text19.Text)
End Sub

Private Sub Text2_Change()
Text15.Text = 1
If Text2.Text = "" Then
Text1.Text = ""
Text3.Text = ""


LimparUltimaescolha
If completarFinalizar = 0 Then
Text5.Text = ""
Text9.Text = ""
End If
End If
End Sub

Private Sub Text2_DblClick()
Text2.Text = ""
End Sub

Private Sub Text2_GotFocus()
'MsgBox Pesquisacardapiofun
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Then
Text3.Text = ""
Text3.SetFocus
End If
If KeyCode = 37 Then
Text1.Text = ""
Text1.SetFocus
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If Text2.Text <> "" Then
KeyAscii = (SoNumeros(KeyAscii))
 If KeyAscii = 13 Then
limparbasedecontas
   KeyAscii = 0

   Animation1.Open Text11.Text

Animation1.AutoPlay = True
funCardapioPorCodigoDoSistema (Text2.Text)
'Text1.Text = ""
'Text3.Text = ""
 Call Command13_Click
End If
Else
MsgBox "Digite o código do produto", , "Código ?"
Text2.SetFocus
End If

End Sub

Public Sub funcall(valor As Double)
Dim meioValor As Double
meioValor = valor / 2




If Text7.Visible = True Then
        'If Text4.Visible = False Then
        'gloValor1 = meioValor
        If Text9.Text <> "" And Text4.Visible = True Then
        Text4.Text = Format(meioValor, "currency")
        Else
        Text9.Text = Format(meioValor, "currency")
       End If
        
        contadorFunCall = True
            If completarFinalizar = 0.5 And Text7.Text <> "" Then
Text7.Text = Text7.Text & "   1/2 " & Text3.Text
Else
 Text7.Text = "1/2  " & Text3.Text
End If
        'Else
        'gloValor2 = meioValor
        'Text4.Text = Format(gloValor1, "currency")
        'Text7.Text = Text7.Text & "  1/2 " & Text3.Text
        'finalizar montagem
        'End If
Else
  
  

        Text9.Text = Format(valor, "currency")
        Text7.Text = " "
       

End If
End Sub



Private Sub Text20_Change()
Text20 = DBLcurre(Text20.Text)
End Sub

Private Sub Text21_Change()
Text21 = DBLcurre(Text21.Text)
End Sub

Private Sub Text22_Change()
If Text16.Text = "0" And Text16.Visible = True Then
Dim totalletras As Integer
Dim comecApartirDe As Integer
Dim sringComoFica As String
Dim retiraadireita As Integer
retiraadireita = InStr(8, Text22.Text, "AC", 0)
totalletras = Len(Text22.Text)


 Do While retiraadireita > 0
          If retiraadireita <> 0 Then
        comecApartirDe = totalletras - retiraadireita - 4
        
        
        sringComoFica = Left(Text22.Text, comecApartirDe)
        'comecApartirDe = InStr(3, Text22.Text, " AC RETIRAR ACRÉSSIMO", 0)
        'sringComoFica = Left(Text22.Text, comecApartirDe)
        Text22.Text = sringComoFica
        End If
        retiraadireita = InStr(8, Text22.Text, "AC", 0)
        totalletras = Len(Text22.Text)
  Loop
  
                                        
If retiraadireita <> 0 Then
comecApartirDe = totalletras - retiraadireita - 4


sringComoFica = Left(Text22.Text, comecApartirDe)
'comecApartirDe = InStr(3, Text22.Text, " AC RETIRAR ACRÉSSIMO", 0)
'sringComoFica = Left(Text22.Text, comecApartirDe)
Text22.Text = sringComoFica
End If

End If
End Sub

Private Sub Text23_Change()
If Text23 <> 1 Then
numerodopedido = 0
End If
End Sub

Private Sub Text24_Change()
Text24 = DBLcurre(Text24.Text)
End Sub

Private Sub Text26_Change()
If Text26.Text = 0 And Label15 <> "Label15" Then
'CANCELAMENTO (Label15.Caption)
End If
End Sub

Private Sub Text3_Change()
Text15.Text = 1
If Text7.Visible = False And Text5 <> "" Then
Text22.Text = Text3 & " " & Text5
Else
Text22.Text = Text7 & " " & Text5
End If
If Text3.Text = "" Then


cobreOvalormaisAlto
REPARSEvalores
Text1.Text = ""
Text2.Text = ""
LimparUltimaescolha
If completarFinalizar = 0 Then
Text5.Text = ""
Text9.Text = ""
End If
End If
End Sub

Private Sub Text3_DblClick()
Text3.Text = ""

End Sub

Private Sub Text3_GotFocus()
If Text12.Visible = True And Text12.Text = "" And Text31.Text <> 1 Then
Form23.Visible = True
Form23.SetFocus
End If
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 38 Then
Text1.Text = ""
Text1.SetFocus
End If
If KeyCode = 39 Then
Text2.Text = ""
Text2.SetFocus
End If

End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
 'consultar
  If KeyAscii = 13 Then
   KeyAscii = 0
 limparbasedecontas
Animation1.Open Text11.Text

Animation1.AutoPlay = True

FunCardapioPordescricao (Trim(Text3.Text))
Call Command13_Click

End If

  If KeyAscii = 8 Then
   KeyAscii = 0
   Text3.Text = ""
   
gloPassos = 1


End If
Animation1.Close
End Sub

Public Sub refuncall()

If completarFinalizar = 0 Then

If Text7.Visible = True And Text9.Text <> "" And contadorFunCall = False Then
        Text9 = Text9 / 2
        Text9.Text = Format(Text9.Text, "currency")
            If completarFinalizar = 0.5 And Text7.Text <> "" Then
        Text7.Text = Text7.Text & "   1/2 " & Text3.Text
        Else
         Text7.Text = "1/2  " & Text3.Text
        End If
        contadorFunCall = True
        
        ElseIf Text9.Text <> "" And contadorFunCall = True Then
        Text9 = Text9 * 2
        Text9.Text = Format(Text9.Text, "currency")
        Text7.Text = " "
         contadorFunCall = False
        
        End If
Else

    If Text7.Visible = True And Text4.Text <> "" And contadorFunCall = False Then
        Text4 = Text4 / 2
        Text4.Text = Format(Text4.Text, "currency")
            If completarFinalizar = 0.5 And Text7.Text <> "" Then
        Text7.Text = Text7.Text & "   1/2 " & Text3.Text
        Else
         Text7.Text = "1/2  " & Text3.Text
        End If
        contadorFunCall = True
        
        ElseIf Text4.Text <> "" And contadorFunCall = True Then
        Text4 = Text4 * 2
        Text4.Text = Format(Text4.Text, "currency")
        Text7.Text = " "
         contadorFunCall = False
        
        End If


End If





End Sub



Public Sub PuraBuscaPorDescrição(Mcodigo As String)
Dim TituloGrid As String
Dim intMcodigo As String
Dim valorObtido As Double
Dim QuantidadeEncontrada As Integer

intMcodigo = Trim("%" + Mcodigo + "%")

ConServer

Dim NpTitulo As String

Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs.ActiveConnection = con
rs.CursorLocation = adUseClient
        


                                     
        Select Case Pesquisacardapiofun
            Case 0 'simples pesquisa
            sql = "SELECT `idCardapio`,`codigo`,`Descricao`,`tipo` FROM `Cardapio` WHERE `Descricao` LIKE'" & intMcodigo & "'ORDER BY `Cardapio`.`Descricao` ASC "
                                                                            rs.Open sql
                                                                             If rs.BOF = False Then
                                                                                    If IsNull(rs.Fields("codigo").Value) Then
                                                                                    Text1.Text = "Não cadastr."
                                                                                    Else
                                                                                     Text1.Text = rs.Fields("codigo").Value
                                                                                     End If
                                                                                    
                                                                              Text2.Text = rs.Fields("idCardapio").Value
                                                                              Text3.Text = rs.Fields("Descricao").Value
                                                                              TituloGrid = rs.Fields("tipo").Value
                                                                            End If
                                                                            rs.Close
                                             
                              'retorno do datagrid pesquisa
                              sql = "SELECT `idCardapio`,`codigo`,`Descricao`,`Medida`,`tipo` ,`valor` FROM `Cardapio` WHERE `Descricao` LIKE'" & intMcodigo & "'ORDER BY `Cardapio`.`Descricao` ASC "
                              rs.Open sql
                             If rs.BOF = False Then
                             Set DataGrid1.DataSource = rs
                             DataGrid1.Visible = True
                             DataGrid1.SetFocus
                        
                             Else
                             MsgBox " O produto : '" & Mcodigo & "'não foi localizado no sistema ", , "Não Existe"
                             DataGrid1.Visible = False
                             
                             End If
                            
                            
                            'rs.Close
                            
                            'DataGrid1.Columns(0).Caption = TituloGrid
                            'DataGrid1.Columns(0).Width = 12000

 Set rs = Nothing
              
            Case 1 'pesquisa por titulos

'conte possibilidades do  produto
sql = "SELECT  COUNT(*)  FROM `Cardapio` WHERE `Descricao` LIKE'" & intMcodigo & "'AND`Titulo` LIKE'" & Text8.Text & "'ORDER BY `Cardapio`.`Descricao` ASC "
                                                                            rs.Open sql
 
QuantidadeEncontrada = rs.Fields("COUNT(*)").Value
            
            rs.Close
            
sql = "SELECT `idCardapio`,`codigo`,`Descricao`,`tipo`,`Medida` ,`valor`  FROM `Cardapio` WHERE `Descricao` LIKE'" & intMcodigo & "'AND`Titulo` LIKE'" & Text8.Text & "'ORDER BY `Cardapio`.`Descricao` ASC "
                                                                            rs.Open sql
                                                                             If rs.BOF = False Then
                                                                             If IsNull(rs.Fields("codigo").Value) Then
                                                                             Text1.Text = "Não cadastr."
                                                                             Else
                                                                              Text1.Text = rs.Fields("codigo").Value
                                                                              End If
                                                                              Text2.Text = rs.Fields("idCardapio").Value
                                                                              Text3.Text = rs.Fields("Descricao").Value
                                                                              TituloGrid = rs.Fields("tipo").Value
                                                                              If QuantidadeEncontrada = 1 Then
                                                                              Text5.Text = rs.Fields("Medida").Value
                                                                              Text9.Text = Format(rs.Fields("valor").Value, "Currency")
                                                                              
                                                                              
                                                                              End If
                                                                              
                                                                              
                                                                            End If
                                                                            rs.Close
                                             
                              
                              sql = "SELECT `idCardapio`,`codigo`,`Descricao`,`Medida`,`tipo` ,`valor` FROM `Cardapio` WHERE `Descricao` LIKE'" & intMcodigo & "'AND`Titulo` LIKE'" & Text8.Text & "'ORDER BY `Cardapio`.`Descricao` ASC "
                              rs.Open sql
                             If rs.BOF = False Then
                             Set DataGrid1.DataSource = rs
                             DataGrid1.Visible = True
                             DataGrid1.SetFocus
                        
                             Else
                             MsgBox " O produto : '" & Mcodigo & "'não foi localizado no sistema ", , "Não Existe"
                             DataGrid1.Visible = False
                             
                             End If
                            
                            
                            'rs.Close
                            
                            'DataGrid1.Columns(0).Caption = TituloGrid
                            'DataGrid1.Columns(0).Width = 12000

 Set rs = Nothing
               
            Case 2 'pesquisar subtitulos
            
'conte possibilidades do  produto
sql = "SELECT  COUNT(*) FROM `Cardapio` WHERE `Descricao` LIKE'" & intMcodigo & "'AND `tipo` LIKE'" & Text8.Text & "'ORDER BY `Cardapio`.`Descricao` ASC "
                                                                            rs.Open sql
 
QuantidadeEncontrada = rs.Fields("COUNT(*)").Value
            
            rs.Close
            
    
              sql = "SELECT `idCardapio`,`codigo`,`Descricao`,`tipo` ,`Medida`,`valor`  FROM `Cardapio` WHERE `Descricao` LIKE'" & intMcodigo & "'AND `tipo` LIKE'" & Text8.Text & "'ORDER BY `Cardapio`.`Descricao` ASC "
                                                                            rs.Open sql
                                                                             If rs.BOF = False Then
                                                                                    If IsNull(rs.Fields("codigo").Value) Then
                                                                                    Text1.Text = "Não cadastr."
                                                                                    Else
                                                                                     Text1.Text = rs.Fields("codigo").Value
                                                                                     End If
                                                                              Text2.Text = rs.Fields("idCardapio").Value
                                                                              Text3.Text = rs.Fields("Descricao").Value
                                                                              TituloGrid = rs.Fields("tipo").Value
                                                                                    If QuantidadeEncontrada = 1 Then
                                                                                    Text5.Text = rs.Fields("Medida").Value
                                                                                    Text9.Text = Format(rs.Fields("valor").Value, "Currency")
                                                                                    End If
                                                                            End If
                                                                            rs.Close
                                             
                              
                              sql = "SELECT `idCardapio`,`codigo`,`Descricao`,`Medida`,`tipo` ,`valor` FROM `Cardapio` WHERE `Descricao` LIKE'" & intMcodigo & "'AND`tipo` LIKE'" & Text8.Text & "'ORDER BY `Cardapio`.`Descricao` ASC "
                              rs.Open sql
                             If rs.BOF = False Then
                             Set DataGrid1.DataSource = rs
                             DataGrid1.Visible = True
                             DataGrid1.SetFocus
                        
                             Else
                             MsgBox " O produto : '" & Mcodigo & "'não foi localizado no sistema , cardápio selecionado indice " & Text8.Text & "", , "Não Existe"
                             DataGrid1.Visible = False
                             
                             End If
                            
                            
                            'rs.Close
                            
                            'DataGrid1.Columns(0).Caption = TituloGrid
                            'DataGrid1.Columns(0).Width = 12000

 Set rs = Nothing
              
              
        End Select
      'DataGrid1.Columns(0).Caption = TituloGrid
                            'DataGrid1.Columns(0).Width = 12000
                                                                    
End Sub



Public Sub buscaDescricaoComCriterios(Mcodigo As String)

Dim TituloGrid As String
Dim intMcodigo As String
Dim valorObtido As Double
Dim Medida As String
intMcodigo = Trim("%" + Text3.Text + "%")
Medida = Trim("%" + Text5.Text + "%")
ConServer

Dim NpTitulo As String
Dim tMedida As String
Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs.ActiveConnection = con
rs.CursorLocation = adUseClient


Dim contagemDePordutos As Integer

                             Select Case Pesquisacardapiofun
            Case 0 'criterios de medidas
           
                      
                     sql = "SELECT COUNT(*) FROM `Cardapio` WHERE `Descricao` LIKE'" & intMcodigo & "' AND`Medida` LIKE'" & Medida & "' ORDER BY `Cardapio`.`Descricao` ASC "
                     rs.Open sql
                     
                      contagemDePordutos = rs.Fields("COUNT(*)").Value
                       If (contagemDePordutos > 1) Then
                       DataGrid1.Visible = True
                       'verificar caso so por um tentando reduzir as possibiilidades
                       rs.Close
                       
                        intMcodigo = Trim(Text3.Text)
                         sql = "SELECT COUNT(*) FROM `Cardapio` WHERE `Descricao` LIKE'" & intMcodigo & "' AND`Medida` LIKE'" & Medida & "' ORDER BY `Cardapio`.`Descricao` ASC "
                        rs.Open sql
                                     If rs.BOF = False Then
                                     If (rs.Fields("COUNT(*)").Value <> 0) Then
                                     contagemDePordutos = rs.Fields("COUNT(*)").Value
                                     Else
                                     intMcodigo = Trim("%" + Text3.Text + "%")
                                      Medida = Trim("%" + Text5.Text + "%")
                                     End If
                                     End If
                       Else 'nova tentativa
                       If rs.BOF = True Then
                                               tMedida = Trim(Text5.Text)
                                            ' tMedida = Trim(Text5.Text)
                                            intMcodigo = Trim(Text3.Text)
                                              rs.Close
                                              
                                               sql = "SELECT COUNT(*) FROM `Cardapio` WHERE `Descricao` LIKE'" & intMcodigo & "' AND`Medida` LIKE'" & Medida & "' ORDER BY `Cardapio`.`Descricao` ASC "
                                            rs.Open sql
                                            
                                             contagemDePordutos = rs.Fields("COUNT(*)").Value
                                              If (contagemDePordutos > 1) Then
                                              DataGrid1.Visible = True
                                              'verificar caso so por um tentando reduzir as possibiilidades
                                              rs.Close
                                                sql = "SELECT COUNT(*) FROM `Cardapio` WHERE `Descricao` LIKE'" & intMcodigo & "' AND`Medida` LIKE'" & Medida & "' ORDER BY `Cardapio`.`Descricao` ASC "
                                               rs.Open sql
                                                            If rs.BOF = False Then
                                                            If (rs.Fields("COUNT(*)").Value <> 0) Then
                                                            contagemDePordutos = rs.Fields("COUNT(*)").Value
                                                           
                                                            End If
                                                            End If
                                End If
                        End If
                       
                                     
                       End If
                        
                       If (contagemDePordutos > 1) Then
                       DataGrid1.Visible = True
                       'DataGrid1.Columns(0).Caption = " POSSÍVEIS PRODUTOS COM  " & Text3.Text
                       'DataGrid1.Columns(0).Width = 12000
                       Else
                       DataGrid1.Visible = False
                       
                       End If
                       
                       
                       
                       
                      
                      If (contagemDePordutos > 1 And DataGrid1.Visible = True) Then
                      rs.Close
                      
                       sql = "SELECT `idCardapio`,`codigo`,`Descricao`,`Medida`,`tipo` ,`valor` FROM `Cardapio` WHERE `Descricao` LIKE'" & intMcodigo & "'ORDER BY `Cardapio`.`Descricao` ASC "
                              rs.Open sql
                             If rs.BOF = False Then
                             Set DataGrid1.DataSource = rs
                             DataGrid1.Visible = True
                             DataGrid1.SetFocus
                             Else
                             MsgBox " O produto : '" & Mcodigo & "'não foi localizado no sistema ", , "Não Existe"
                             DataGrid1.Visible = False
                             
                             End If
                      Else
                      
                
                                                                             rs.Close
                                                                            sql = "SELECT `idCardapio`,`codigo`,`Descricao`,`tipo`,`valor`,`Medida` FROM `Cardapio` WHERE `Descricao` LIKE'" & intMcodigo & "' AND`Medida` LIKE'" & Medida & "'ORDER BY `Cardapio`.`Descricao` ASC"
                                                                           ' sql = "SELECT `idCardapio`,`codigo`,`Descricao`,`tipo`,`valor`,`Medida` FROM `Cardapio` WHERE `Descricao` LIKE AND`Medida` LIKE ORDER BY `Cardapio`.`Descricao` ASC "
                                                                            rs.Open sql
                                                                             If rs.BOF = False Then
                                                                                    If IsNull(rs.Fields("codigo").Value) Then
                                                                                    Text1.Text = "Não cadastr."
                                                                                    Else
                                                                                    Text1.Text = rs.Fields("codigo").Value
                                                                                     End If
                                                                              Text2.Text = rs.Fields("idCardapio").Value
                                                                              Text3.Text = rs.Fields("Descricao").Value
                                                                              'NpTitulo = rs.Fields("tipo").Value
                                                                             funcall (rs.Fields("valor").Value)
                                                                             DataGrid1.Visible = False
                                                                              
                                                                              Else
                                                                              
                                                                              MsgBox "O produto " & Text3.Text & " não esta disponivel na medida compatível"
                                                                            End If
                                                                            rs.Close

                      End If





 Set rs = Nothing
            
            
            Case 1 'critero de medida mais titulo
              
            
              
                     sql = "SELECT COUNT(*) FROM `Cardapio` WHERE `Descricao` LIKE'" & intMcodigo & "' AND`Medida` LIKE'" & Medida & "'  AND`Titulo` LIKE'" & Text8.Text & "'ORDER BY `Cardapio`.`Descricao` ASC "
                     rs.Open sql
                     
                      contagemDePordutos = rs.Fields("COUNT(*)").Value
                       If (contagemDePordutos > 1) Then
                       DataGrid1.Visible = True
                       'verificar caso so por um tentando reduzir as possibiilidades
                       rs.Close
                       
                        intMcodigo = Trim(Text3.Text)
                         sql = "SELECT COUNT(*) FROM `Cardapio` WHERE `Descricao` LIKE'" & intMcodigo & "' AND`Medida` LIKE'" & Medida & "'  AND`Titulo` LIKE'" & Text8.Text & "'ORDER BY `Cardapio`.`Descricao` ASC "
                        rs.Open sql
                                     If rs.BOF = False Then
                                     If (rs.Fields("COUNT(*)").Value <> 0) Then
                                     contagemDePordutos = rs.Fields("COUNT(*)").Value
                                     Else
                                     intMcodigo = Trim("%" + Text3.Text + "%")
                                      Medida = Trim("%" + Text5.Text + "%")
                                     End If
                                     End If
                       Else 'nova tentativa
                       If rs.BOF = True Then
                                               tMedida = Trim(Text5.Text)
                                            ' tMedida = Trim(Text5.Text)
                                            intMcodigo = Trim(Text3.Text)
                                              rs.Close
                                              
                                               sql = "SELECT COUNT(*) FROM `Cardapio` WHERE `Descricao` LIKE'" & intMcodigo & "' AND`Medida` LIKE'" & Medida & "'  AND`Titulo` LIKE'" & Text8.Text & "'ORDER BY `Cardapio`.`Descricao` ASC "
                                            rs.Open sql
                                            
                                             contagemDePordutos = rs.Fields("COUNT(*)").Value
                                              If (contagemDePordutos > 1) Then
                                              DataGrid1.Visible = True
                                              'verificar caso so por um tentando reduzir as possibiilidades
                                              rs.Close
                                                sql = "SELECT COUNT(*) FROM `Cardapio` WHERE `Descricao` LIKE'" & intMcodigo & "' AND`Medida` LIKE'" & Medida & "'  AND`Titulo` LIKE'" & Text8.Text & "'ORDER BY `Cardapio`.`Descricao` ASC "
                                               rs.Open sql
                                                            If rs.BOF = False Then
                                                            If (rs.Fields("COUNT(*)").Value <> 0) Then
                                                            contagemDePordutos = rs.Fields("COUNT(*)").Value
                                                           
                                                            End If
                                                            End If
                                End If
                        End If
                       
                                     
                       End If
                        
                       If (contagemDePordutos > 1) Then
                       DataGrid1.Visible = True
                       'DataGrid1.Columns(0).Caption = " POSSÍVEIS PRODUTOS COM  " & Text3.Text
                       'DataGrid1.Columns(0).Width = 12000
                       Else
                       DataGrid1.Visible = False
                       
                       End If
                       
                       
                       
                       
                      
                      If (contagemDePordutos > 1 And DataGrid1.Visible = True) Then
                      rs.Close
                      
                       sql = "SELECT `idCardapio`,`codigo`,`Descricao`,`Medida`,`tipo` ,`valor` FROM `Cardapio` WHERE `Descricao` LIKE'" & intMcodigo & "' AND`Titulo` LIKE'" & Text8.Text & "'ORDER BY `Cardapio`.`Descricao` ASC "
                              rs.Open sql
                             If rs.BOF = False Then
                             Set DataGrid1.DataSource = rs
                             DataGrid1.Visible = True
                             DataGrid1.SetFocus
                             Else
                             MsgBox " O produto : '" & Mcodigo & "'não foi localizado no sistema ", , "Não Existe"
                             DataGrid1.Visible = False
                             
                             End If
                      Else
                      
                
                                                                             rs.Close
                                                                            sql = "SELECT `idCardapio`,`codigo`,`Descricao`,`tipo`,`valor`,`Medida` FROM `Cardapio` WHERE `Descricao` LIKE'" & intMcodigo & "' AND`Medida` LIKE'" & Medida & "' AND`Titulo` LIKE'" & Text8.Text & "'ORDER BY `Cardapio`.`Descricao` ASC"
                                                                           ' sql = "SELECT `idCardapio`,`codigo`,`Descricao`,`tipo`,`valor`,`Medida` FROM `Cardapio` WHERE `Descricao` LIKE AND`Medida` LIKE ORDER BY `Cardapio`.`Descricao` ASC "
                                                                            rs.Open sql
                                                                             If rs.BOF = False Then
                                                                             
                                                                                    If IsNull(rs.Fields("codigo").Value) Then
                                                                                    Text1.Text = "Não cadastr."
                                                                                    Else
                                                                                    Text1.Text = rs.Fields("codigo").Value
                                                                                     End If
                                                                              Text2.Text = rs.Fields("idCardapio").Value
                                                                              Text3.Text = rs.Fields("Descricao").Value
                                                                              'NpTitulo = rs.Fields("tipo").Value
                                                                             funcall (rs.Fields("valor").Value)
                                                                             DataGrid1.Visible = False
                                                                              
                                                                              Else
                                                                              
                                                                              MsgBox "O produto " & Text3.Text & " não esta disponivel na medida compatível"
                                                                            End If
                                                                            rs.Close

                      End If





 Set rs = Nothing
            
            
            Case 2 'criterio de medida subtitulo
            
                     sql = "SELECT COUNT(*) FROM `Cardapio` WHERE `Descricao` LIKE'" & intMcodigo & "' AND`Medida` LIKE'" & Medida & "'  AND`tipo` LIKE'" & Text8.Text & "'ORDER BY `Cardapio`.`Descricao` ASC "
                     rs.Open sql
                     
                      contagemDePordutos = rs.Fields("COUNT(*)").Value
                       If (contagemDePordutos > 1) Then
                       DataGrid1.Visible = True
                       'verificar caso so por um tentando reduzir as possibiilidades
                       rs.Close
                       
                        intMcodigo = Trim(Text3.Text)
                         sql = "SELECT COUNT(*) FROM `Cardapio` WHERE `Descricao` LIKE'" & intMcodigo & "' AND`Medida` LIKE'" & Medida & "'  AND`tipo` LIKE'" & Text8.Text & "'ORDER BY `Cardapio`.`Descricao` ASC "
                        rs.Open sql
                                     If rs.BOF = False Then
                                     If (rs.Fields("COUNT(*)").Value <> 0) Then
                                     contagemDePordutos = rs.Fields("COUNT(*)").Value
                                     Else
                                     intMcodigo = Trim("%" + Text3.Text + "%")
                                      Medida = Trim("%" + Text5.Text + "%")
                                     End If
                                     End If
                       Else 'nova tentativa
                       If rs.BOF = True Then
                                               tMedida = Trim(Text5.Text)
                                            ' tMedida = Trim(Text5.Text)
                                            intMcodigo = Trim(Text3.Text)
                                              rs.Close
                                              
                                               sql = "SELECT COUNT(*) FROM `Cardapio` WHERE `Descricao` LIKE'" & intMcodigo & "' AND`Medida` LIKE'" & Medida & "'  AND`tipo` LIKE'" & Text8.Text & "'ORDER BY `Cardapio`.`Descricao` ASC "
                                            rs.Open sql
                                            
                                             contagemDePordutos = rs.Fields("COUNT(*)").Value
                                              If (contagemDePordutos > 1) Then
                                              DataGrid1.Visible = True
                                              'verificar caso so por um tentando reduzir as possibiilidades
                                              rs.Close
                                                sql = "SELECT COUNT(*) FROM `Cardapio` WHERE `Descricao` LIKE'" & intMcodigo & "' AND`Medida` LIKE'" & Medida & "' AND`tipo` LIKE'" & Text8.Text & "' ORDER BY `Cardapio`.`Descricao` ASC "
                                               rs.Open sql
                                                            If rs.BOF = False Then
                                                            If (rs.Fields("COUNT(*)").Value <> 0) Then
                                                            contagemDePordutos = rs.Fields("COUNT(*)").Value
                                                           
                                                            End If
                                                            End If
                                End If
                        End If
                       
                                     
                       End If
                        
                       If (contagemDePordutos > 1) Then
                       DataGrid1.Visible = True
                       'DataGrid1.Columns(0).Caption = " POSSÍVEIS PRODUTOS COM  " & Text3.Text
                       'DataGrid1.Columns(0).Width = 12000
                       Else
                       DataGrid1.Visible = False
                       
                       End If
                       
                       
                       
                       
                      
                      If (contagemDePordutos > 1 And DataGrid1.Visible = True) Then
                      rs.Close
                      
                       sql = "SELECT `idCardapio`,`codigo`,`Descricao`,`Medida`,`tipo` ,`valor` FROM `Cardapio` WHERE `Descricao` LIKE'" & intMcodigo & "' AND`tipo` LIKE'" & Text8.Text & "'ORDER BY `Cardapio`.`Descricao` ASC "
                              rs.Open sql
                             If rs.BOF = False Then
                             Set DataGrid1.DataSource = rs
                             DataGrid1.Visible = True
                             DataGrid1.SetFocus
                             Else
                             MsgBox " O produto : '" & Mcodigo & "'não foi localizado no sistema ", , "Não Existe"
                             DataGrid1.Visible = False
                             
                             End If
                      Else
                      
                
                                                                             rs.Close
                                                                            sql = "SELECT `idCardapio`,`codigo`,`Descricao`,`tipo`,`valor`,`Medida` FROM `Cardapio` WHERE `Descricao` LIKE'" & intMcodigo & "' AND`Medida` LIKE'" & Medida & "' AND`tipo` LIKE'" & Text8.Text & "'ORDER BY `Cardapio`.`Descricao` ASC"
                                                                           ' sql = "SELECT `idCardapio`,`codigo`,`Descricao`,`tipo`,`valor`,`Medida` FROM `Cardapio` WHERE `Descricao` LIKE AND`Medida` LIKE ORDER BY `Cardapio`.`Descricao` ASC "
                                                                            rs.Open sql
                                                                             If rs.BOF = False Then
                                                                           
                                                                                    If IsNull(rs.Fields("codigo").Value) Then
                                                                                    Text1.Text = "Não cadastr."
                                                                                    Else
                                                                                    Text1.Text = rs.Fields("codigo").Value
                                                                                     End If
                                                                              Text2.Text = rs.Fields("idCardapio").Value
                                                                              Text3.Text = rs.Fields("Descricao").Value
                                                                              'NpTitulo = rs.Fields("tipo").Value
                                                                             funcall (rs.Fields("valor").Value)
                                                                             DataGrid1.Visible = False
                                                                              
                                                                              Else
                                                                              
                                                                              MsgBox "O produto " & Text3.Text & " não esta disponivel na medida compatível"
                                                                            End If
                                                                            rs.Close

                      End If





 Set rs = Nothing
            
             
            
             
             
             
         
        End Select


'DataGrid1.Columns(0).Caption = " Escolha entre os possíveis produtos"
                            'DataGrid1.Columns(0).Width = 12000



End Sub
Public Function definaAsmedidaspossiveis(Mcodigo As String) As Boolean
'form500.show
Dim UNidadeDisponiveis As Integer
Dim intMcodigo As String
intMcodigo = Trim("%" + Mcodigo + "%")
'Text3.Text = Mcodigo
   Animation1.Open Text11.Text

Form500.ProgressBar1.Value = 20
Animation1.AutoPlay = True
'ConServer

Dim NpTitulo As String
Form500.ProgressBar1.Value = 25
Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset
Form500.ProgressBar1.Value = 30
Set rs.ActiveConnection = con
rs.CursorLocation = adUseClient

Form500.ProgressBar1.Value = 32


Select Case Pesquisacardapiofun
            Case 0 'Medida simples de busca comum
            
            
                                                    
''Sleep 1000
        

          '
       'RETORNOO DE VARIAVAIS
                                        '  sql = "SELECT COUNT (`Medida`) FROM `Cardapio` WHERE `Descricao` LIKE "
                                                
                                                sql = "SELECT COUNT(DISTINCT `Medida`)  FROM `Cardapio` WHERE `Descricao` LIKE '" & intMcodigo & "'ORDER BY `Medida` ASC"
                                                
                                                rs.Open sql
                                                UNidadeDisponiveis = rs.Fields("COUNT(DISTINCT `Medida`)").Value
                                                Form500.ProgressBar1.Value = 35
                                                            If UNidadeDisponiveis < 1 Then
                                                            rs.Close
                                                              sql = "SELECT COUNT(DISTINCT `Medida`)  FROM `Cardapio` WHERE `Descricao` LIKE '" & Trim(Text3.Text) & "'ORDER BY `Medida` ASC"
                                                                 rs.Open sql
                                                                     UNidadeDisponiveis = rs.Fields("COUNT(DISTINCT `Medida`)").Value
                                                                      If UNidadeDisponiveis < 1 Then
                                                                         intMcodigo = Trim(Mcodigo)
                                                                            rs.Close
                                                                                sql = "SELECT COUNT(DISTINCT `Medida`)  FROM `Cardapio` WHERE `Descricao` LIKE '" & intMcodigo & "'ORDER BY `Medida` ASC"
                                                                                 rs.Open sql
                                                                                     UNidadeDisponiveis = rs.Fields("COUNT(DISTINCT `Medida`)").Value
                                                                                     
                                                                                      If UNidadeDisponiveis < 1 Then
                                                                         intMcodigo = Mcodigo
                                                                            rs.Close
                                                                                sql = "SELECT COUNT(DISTINCT `Medida`)  FROM `Cardapio` WHERE `Descricao` LIKE '" & intMcodigo & "'ORDER BY `Medida` ASC"
                                                                                 rs.Open sql
                                                                                     UNidadeDisponiveis = rs.Fields("COUNT(DISTINCT `Medida`)").Value
                                                                                     End If
                                                                                     
                                                                                     
                                                                                     
                                                                                     
                                                                                     End If
                                                                     
                                                                     
                                                                     
                                                             End If
                                            
                                                            Form500.ProgressBar1.Value = 50
                                                           
                                                
                                                                                    
                                                rs.Close
                                                         'SELECT DISTINCT `Medida` FROM `Cardapio` WHERE `Descricao` LIKE 'BACON E MILHO'ORDER BY `Cardapio`.`Medida` ASC
                                                sql = "SELECT DISTINCT `Medida` FROM `Cardapio` WHERE `Descricao` LIKE '" & intMcodigo & "' ORDER BY `Cardapio`.`Medida` ASC"
                                                rs.Open sql
                                                 If rs.BOF = False And UNidadeDisponiveis >= 1 Then
                                                 If UNidadeDisponiveis = 1 Then
                                                 ' caso item somente uma unidade logica indivisivel?
                                                 
                                                 End If
                                                  Form500.ProgressBar1.Value = 56
                                                  definaAsmedidaspossiveis = True
                                                  Set DataGrid1.DataSource = rs
                                                  DataGrid1.Caption = "Escolha entre as medidas disponíveis"
                                                  DataGrid1.Visible = True
                                                 
                                                  DataGrid1.SetFocus
                                                 ' mobilidadeMedida = 1
                                               Else
               
  definaAsmedidaspossiveis = False


 MsgBox " O produto  : '" & Mcodigo & "' sem medidas disponiveis  ", , "Não existe medida para este produto"
 
 DataGrid1.Visible = False
 
 End If
Form500.ProgressBar1.Value = 85

'Text3.Text = (DataGrid1.Columns(0).Value)

If definaAsmedidaspossiveis = True Then
'DataGrid1.Columns(0).Caption = "MEDIDAS POSSÍVEIS "
'DataGrid1.Columns(0).Width = 12000
Else
'gloPassos = 1
'organizadoDePassos (gloPassos)
End If
'contadorDePassos = contadorDePassos + 1
Animation1.Close
Form500.ProgressBar1.Value = 100
 Set rs = Nothing
'__________________________________________________________________________________________




          
            Case 1 'pesquisa por titulos
                                                               'sql = "SELECT `idCardapio`,`codigo`,`Descricao` ,`tipo`,`valor`,`Medida`FROM `Cardapio` WHERE `idCardapio` = '" & intMcodigo & "'AND`Titulo` LIKE'" & Text8.Text & "'ORDER BY `Cardapio`.`Descricao` ASC "
                                             '                                               rs.Open sql
                                            '
                                            
                                            ''Sleep 1000
                                                    
                                            
                                                      '
                                                   'RETORNOO DE VARIAVAIS
                                                                                    '  sql = "SELECT COUNT (`Medida`) FROM `Cardapio` WHERE `Descricao` LIKE "
                                                                                            Form500.ProgressBar1.Value = 35
                                                                                            sql = "SELECT COUNT(DISTINCT `Medida`)  FROM `Cardapio` WHERE `Descricao` LIKE '" & intMcodigo & "'AND`Titulo` LIKE'" & Text8.Text & "'ORDER BY  `Medida` ASC"
                                                                                            
                                                                                            rs.Open sql
                                                                                            UNidadeDisponiveis = rs.Fields("COUNT(DISTINCT `Medida`)").Value
                                                                                            
                                                                                                        If UNidadeDisponiveis < 1 Then
                                                                                                        rs.Close
                                                                                                          sql = "SELECT COUNT(DISTINCT `Medida`)  FROM `Cardapio` WHERE `Descricao` LIKE '" & Trim(Text3.Text) & "'AND`Titulo` LIKE'" & Text8.Text & "'ORDER BY  `Medida` ASC"
                                                                                                             rs.Open sql
                                                                                                                 UNidadeDisponiveis = rs.Fields("COUNT(DISTINCT `Medida`)").Value
                                                                                                                  If UNidadeDisponiveis < 1 Then
                                                                                                                     intMcodigo = Trim(Mcodigo)
                                                                                                                        rs.Close
                                                                                                                            sql = "SELECT COUNT(DISTINCT `Medida`)  FROM `Cardapio` WHERE `Descricao` LIKE '" & intMcodigo & "'AND`Titulo` LIKE'" & Text8.Text & "'ORDER BY  `Medida` ASC"
                                                                                                                             rs.Open sql
                                                                                                                                 UNidadeDisponiveis = rs.Fields("COUNT(DISTINCT `Medida`)").Value
                                                                                                                                 
                                                                                                                                  If UNidadeDisponiveis < 1 Then
                                                                                                                     intMcodigo = Mcodigo
                                                                                                                        rs.Close
                                                                                                                            sql = "SELECT COUNT(DISTINCT `Medida`)  FROM `Cardapio` WHERE `Descricao` LIKE '" & intMcodigo & "'AND`Titulo` LIKE'" & Text8.Text & "'ORDER BY  `Medida` ASC"
                                                                                                                             rs.Open sql
                                                                                                                                 UNidadeDisponiveis = rs.Fields("COUNT(DISTINCT `Medida`)").Value
                                                                                                                                 End If
                                                                                                                                 
                                                                                                                                 
                                                                                                                                 
                                                                                                                                 
                                                                                                                                 End If
                                                                                                                 
                                                                                                                 
                                                                                                                 
                                                                                                         End If
                                                                                        
                                                                                                        
                                                                                                       
                                                                                            Form500.ProgressBar1.Value = 40
                                                                                                                                
                                                                                            rs.Close
                                                                                                     'SELECT DISTINCT `Medida` FROM `Cardapio` WHERE `Descricao` LIKE 'BACON E MILHO'AND`Titulo` LIKE'" & Text8.Text & "'ORDER BY  `Cardapio`.`Medida` ASC
                                                                                            sql = "SELECT DISTINCT `Medida` FROM `Cardapio` WHERE `Descricao` LIKE '" & intMcodigo & "' AND`Titulo` LIKE'" & Text8.Text & "'ORDER BY `Cardapio`.`Medida` ASC"
                                                                                            rs.Open sql
                                                                                             If rs.BOF = False And UNidadeDisponiveis >= 1 Then
                                                                                             If UNidadeDisponiveis = 1 Then
                                                                                             ' caso item somente uma unidade logica indivisivel?
                                                                                             
                                                                                             End If
                                                                                              Form500.ProgressBar1.Value = 45
                                                                                              definaAsmedidaspossiveis = True
                                                                                              Set DataGrid1.DataSource = rs
                                                                                              DataGrid1.Caption = "Escolha entre as medidas disponíveis"
                                                                                              DataGrid1.Visible = True
                                                                                             
                                                                                              DataGrid1.SetFocus
                                                                                             ' mobilidadeMedida = 1
                                                                                           Else
                                                           
                                              definaAsmedidaspossiveis = False
                                            
                                            
                                             MsgBox " O produto  : '" & Mcodigo & "' sem medidas disponiveis  ", , "Não existe medida para este produto"
                                             
                                             DataGrid1.Visible = False
                                             
                                             End If
                                            
                                            Form500.ProgressBar1.Value = 70
                                            'Text3.Text = (DataGrid1.Columns(0).Value)
                                            
                                            If definaAsmedidaspossiveis = True Then
                                            'DataGrid1.Columns(0).Caption = "MEDIDAS POSSÍVEIS "
                                            'DataGrid1.Columns(0).Width = 12000
                                            Else
                                            'gloPassos = 1
                                            'organizadoDePassos (gloPassos)
                                            End If
                                            'contadorDePassos = contadorDePassos + 1
                                            Animation1.Close
                                            Form500.ProgressBar1.Value = 100
                                            Set rs = Nothing
                                            'mostraasmedidas = False
                                                                   

            Case 2 'medidas por pesquisar subtitulos
            


            
            
                                                    
''Sleep 1000
        

          '
       'RETORNOO DE VARIAVAIS
                                        '  sql = "SELECT COUNT (`Medida`) FROM `Cardapio` WHERE `Descricao` LIKE "
                                                Form500.ProgressBar1.Value = 45
                                                sql = "SELECT COUNT(DISTINCT `Medida`)  FROM `Cardapio` WHERE `Descricao` LIKE '" & intMcodigo & "'AND`tipo` LIKE'" & Text8.Text & "'ORDER BY `Medida` ASC"
                                                
                                                rs.Open sql
                                                UNidadeDisponiveis = rs.Fields("COUNT(DISTINCT `Medida`)").Value
                                                
                                                            If UNidadeDisponiveis < 1 Then
                                                            rs.Close
                                                              sql = "SELECT COUNT(DISTINCT `Medida`)  FROM `Cardapio` WHERE `Descricao` LIKE '" & Trim(Text3.Text) & "'AND`tipo` LIKE'" & Text8.Text & "'ORDER BY `Medida` ASC"
                                                                 rs.Open sql
                                                                     UNidadeDisponiveis = rs.Fields("COUNT(DISTINCT `Medida`)").Value
                                                                      If UNidadeDisponiveis < 1 Then
                                                                         intMcodigo = Trim(Mcodigo)
                                                                            rs.Close
                                                                                sql = "SELECT COUNT(DISTINCT `Medida`)  FROM `Cardapio` WHERE `Descricao` LIKE '" & intMcodigo & "'AND`tipo` LIKE'" & Text8.Text & "'ORDER BY `Medida` ASC"
                                                                                 rs.Open sql
                                                                                     UNidadeDisponiveis = rs.Fields("COUNT(DISTINCT `Medida`)").Value
                                                                                     
                                                                                      If UNidadeDisponiveis < 1 Then
                                                                         intMcodigo = Mcodigo
                                                                            rs.Close
                                                                                sql = "SELECT COUNT(DISTINCT `Medida`)  FROM `Cardapio` WHERE `Descricao` LIKE '" & intMcodigo & "'ORDER BY `Medida` ASC"
                                                                                 rs.Open sql
                                                                                     UNidadeDisponiveis = rs.Fields("COUNT(DISTINCT `Medida`)").Value
                                                                                     End If
                                                                                     
                                                                                     
                                                                                     
                                                                                     
                                                                                     End If
                                                                     
                                                                     
                                                                     
                                                             End If
                                            
                                                            
                                                           
                                                Form500.ProgressBar1.Value = 58
                                                                                    
                                                rs.Close
                                                         'SELECT DISTINCT `Medida` FROM `Cardapio` WHERE `Descricao` LIKE 'BACON E MILHO'ORDER BY `Cardapio`.`Medida` ASC
                                                sql = "SELECT DISTINCT `Medida` FROM `Cardapio` WHERE `Descricao` LIKE '" & intMcodigo & "' AND`tipo` LIKE'" & Text8.Text & "'ORDER BY `Cardapio`.`Medida` ASC"
                                                rs.Open sql
                                                 If rs.BOF = False And UNidadeDisponiveis >= 1 Then
                                                 If UNidadeDisponiveis = 1 Then
                                                 ' caso item somente uma unidade logica indivisivel?
                                                 
                                                 End If
                                                  Form500.ProgressBar1.Value = 65
                                                  definaAsmedidaspossiveis = True
                                                  Set DataGrid1.DataSource = rs
                                                  DataGrid1.Caption = "Escolha entre as medidas disponíveis"
                                                  DataGrid1.Visible = True
                                                 
                                                  DataGrid1.SetFocus
                                                 ' mobilidadeMedida = 1
                                               Else
               Form500.ProgressBar1.Value = 70
  definaAsmedidaspossiveis = False


 MsgBox " O produto  : '" & Mcodigo & "' sem medidas disponiveis  ", , "Não existe medida para este produto"
 
 DataGrid1.Visible = False
 
 End If
Form500.ProgressBar1.Value = 75

'Text3.Text = (DataGrid1.Columns(0).Value)

If definaAsmedidaspossiveis = True Then
'DataGrid1.Columns(0).Caption = "MEDIDAS POSSÍVEIS "
'DataGrid1.Columns(0).Width = 12000
Else
'gloPassos = 1
'organizadoDePassos (gloPassos)
End If
Form500.ProgressBar1.Value = 80
'contadorDePassos = contadorDePassos + 1
Animation1.Close
 Set rs = Nothing
'__________________________________________________________________________________________
Form500.ProgressBar1.Value = 100






        End Select




End Function

Public Function FunCardapioPordescricao(Mcodigo As String) As Boolean



If Text5 = "" Then
'busca comum
PuraBuscaPorDescrição (Trim(Text3.Text))
FunCardapioPordescricao = True

Else
' busque por criterios
buscaDescricaoComCriterios (Trim(Text3.Text))
FunCardapioPordescricao = False
End If

                            


End Function

Public Sub saberOcodigo()
'essa funcão permite sabeer o codigo do produto
Dim valorObtido As Double

Dim nomeDoProduto As String
Dim tamanhoDoProduto As String
nomeDoProduto = ("%" + Text3.Text + "%")
tamanhoDoProduto = ("%" + DataGrid1.Columns(0).Value + "%")

'ConServer



Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs.ActiveConnection = con
'rs.CursorLocation = adUseClient
      
      
      
      Select Case Pesquisacardapiofun
            Case 0 'simples pesquisa

                                                     'SELECT `idCardapio` FROM `Cardapio` WHERE `Descricao` LIKE 'BACON E MILHO' AND `Medida` LIKE 'grande'
                                               sql = "SELECT `idCardapio` ,`codigo` ,`valor`FROM `Cardapio` WHERE `Descricao` LIKE '" & nomeDoProduto & "' AND `Medida` LIKE '" & tamanhoDoProduto & "'"
                                                rs.Open sql
                                                If rs.BOF = True Then
                                                  rs.Close
                                                  sql = "SELECT `idCardapio` ,`codigo` ,`valor`FROM `Cardapio` WHERE `Descricao` LIKE '" & Text3.Text & "' AND `Medida` LIKE '" & DataGrid1.Columns(0).Value & "'"
                                                 rs.Open sql
                                                    If rs.BOF = True Then
                                                  rs.Close
                                                  sql = "SELECT `idCardapio` ,`codigo` ,`valor`FROM `Cardapio` WHERE `Descricao` LIKE '" & Trim(Text3.Text) & "' AND `Medida` LIKE '" & DataGrid1.Columns(0).Value & "'"
                                                 rs.Open sql
                                                     If rs.BOF = True Then
                                                  rs.Close
                                                  sql = "SELECT `idCardapio` ,`codigo` ,`valor`FROM `Cardapio` WHERE `Descricao` LIKE '" & Trim(Text3.Text) & "' AND `Medida` LIKE '" & Text5.Text & "'"
                                                 rs.Open sql
                                                   If rs.BOF = True Then
                                                  rs.Close
                                                  sql = "SELECT `idCardapio` ,`codigo` ,`valor`FROM `Cardapio` WHERE `Descricao` LIKE '" & nomeDoProduto & "' AND `Medida` LIKE '" & Text5.Text & "'"
                                                 rs.Open sql
                                                
                                                  If rs.BOF = True Then
                                                  rs.Close
                                                  sql = "SELECT `idCardapio` ,`codigo` ,`valor`FROM `Cardapio` WHERE `Descricao` LIKE '" & nomeDoProduto & "' AND `Medida` LIKE '" & Text5.Text & "'"
                                                 rs.Open sql
                                                
                                                
                                                
                                                End If
                                                
                                                End If
                                                
                                                
                                                End If
                                                
                                                
                                                End If
                                                
                                                
                                                End If
                                                
                                                
                                                
                                                
                                                If rs.BOF = False Then
                                                'erro de codigo tamanho obturado
                                                  Text5.Text = DataGrid1.Columns(0).Value
                                                  Text2.Text = rs.Fields("idCardapio").Value
                                                  If IsNull(rs.Fields("codigo").Value) Then
                                                  Text1.Text = "não"
                                                  'reparar
                                                  'Exit Sub
                                                  Else
                                                  Text1.Text = rs.Fields("codigo").Value
                                                       End If
                                                        funcall (rs.Fields("valor").Value)
                                                                    
                                                  End If
                                                  
                                                  
                                                 'seguir ou nao para o proximo passo
                                                 If Text5.Text <> DataGrid1.Columns(0).Value Then
                                                  'contadorDePassos = contadorDePassos - 1
                                                 
                                                 
                                                 End If
                                                
                                                rs.Close
         
  
 

 'DataGrid1.Visible = F
 


 Set rs = Nothing

            

               Set rs = Nothing
              
            Case 1 'pesquisa por titulos
'sql = "SELECT `idCardapio`,`codigo`,`Descricao` ,`tipo`,`valor`,`Medida`FROM `Cardapio` WHERE `idCardapio` = '" & intMcodigo & "'AND`Titulo` LIKE'" & Text8.Text & "'ORDER BY `Cardapio`.`Descricao` ASC "
'                                                rs.Open sql
                                                 

                                                     'SELECT `idCardapio` FROM `Cardapio` WHERE `Descricao` LIKE 'BACON E MILHO' AND `Medida` LIKE 'grande'
                                               sql = "SELECT `idCardapio` ,`codigo` ,`valor`FROM `Cardapio` WHERE `Descricao` LIKE '" & nomeDoProduto & "' AND `Medida` LIKE '" & tamanhoDoProduto & "'AND`Titulo` LIKE'" & Text8.Text & "'ORDER BY `Cardapio`.`Descricao` ASC "
                                                rs.Open sql
                                                If rs.BOF = True Then
                                                  rs.Close
                                                  sql = "SELECT `idCardapio` ,`codigo` ,`valor`FROM `Cardapio` WHERE `Descricao` LIKE '" & Text3.Text & "' AND `Medida` LIKE '" & DataGrid1.Columns(0).Value & "'AND`Titulo` LIKE'" & Text8.Text & "'ORDER BY `Cardapio`.`Descricao` ASC "
                                                 rs.Open sql
                                                    If rs.BOF = True Then
                                                  rs.Close
                                                  sql = "SELECT `idCardapio` ,`codigo` ,`valor`FROM `Cardapio` WHERE `Descricao` LIKE '" & Trim(Text3.Text) & "' AND `Medida` LIKE '" & DataGrid1.Columns(0).Value & "'AND`Titulo` LIKE'" & Text8.Text & "'ORDER BY `Cardapio`.`Descricao` ASC "
                                                 rs.Open sql
                                                     If rs.BOF = True Then
                                                  rs.Close
                                                  sql = "SELECT `idCardapio` ,`codigo` ,`valor`FROM `Cardapio` WHERE `Descricao` LIKE '" & Trim(Text3.Text) & "' AND `Medida` LIKE '" & Text5.Text & "' AND`Titulo` LIKE'" & Text8.Text & "'ORDER BY `Cardapio`.`Descricao` ASC "
                                                 rs.Open sql
                                                   If rs.BOF = True Then
                                                  rs.Close
                                                  sql = "SELECT `idCardapio` ,`codigo` ,`valor`FROM `Cardapio` WHERE `Descricao` LIKE '" & nomeDoProduto & "' AND `Medida` LIKE '" & Text5.Text & "'AND`Titulo` LIKE'" & Text8.Text & "'ORDER BY `Cardapio`.`Descricao` ASC "
                                                 rs.Open sql
                                                
                                                  If rs.BOF = True Then
                                                  rs.Close
                                                  sql = "SELECT `idCardapio` ,`codigo` ,`valor`FROM `Cardapio` WHERE `Descricao` LIKE '" & nomeDoProduto & "' AND `Medida` LIKE '" & Text5.Text & "'AND`Titulo` LIKE'" & Text8.Text & "'ORDER BY `Cardapio`.`Descricao` ASC "
                                                 rs.Open sql
                                                
                                                
                                                
                                                End If
                                                
                                                End If
                                                
                                                
                                                End If
                                                
                                                
                                                End If
                                                
                                                
                                                End If
                                                
                                                
                                                
                                                
                                                If rs.BOF = False Then
                                                  Text5.Text = DataGrid1.Columns(0).Value
                                                  Text2.Text = rs.Fields("idCardapio").Value
                                                  If Not IsNull(rs.Fields("codigo").Value) Then
                                                  Text1.Text = rs.Fields("codigo").Value
                                                  End If
                                                        funcall (rs.Fields("valor").Value)
                                                                    
                                                  End If
                                                  
                                                  
                                                 'seguir ou nao para o proximo passo
                                                 If Text5.Text <> DataGrid1.Columns(0).Value Then
                                                  'contadorDePassos = contadorDePassos - 1
                                                 
                                                 
                                                 End If
                                                
                                                rs.Close
         
  
 

 'DataGrid1.Visible = F
 


 Set rs = Nothing


            Case 2 'pesquisar subtitulos
            
'sql = "SELECT `idCardapio`,`codigo`,`Descricao` ,`tipo`,`valor`,`Medida`FROM `Cardapio` WHERE `idCardapio` = '" & intMcodigo & "'AND`tipo` LIKE'" & Text8.Text & "'ORDER BY `Cardapio`.`Descricao` ASC "
 '                                               rs.Open sql
                                                 

                                                     'SELECT `idCardapio` FROM `Cardapio` WHERE `Descricao` LIKE 'BACON E MILHO' AND `Medida` LIKE 'grande'
                                               sql = "SELECT `idCardapio` ,`codigo` ,`valor`FROM `Cardapio` WHERE `Descricao` LIKE '" & nomeDoProduto & "' AND `Medida` LIKE '" & tamanhoDoProduto & "'AND`tipo` LIKE'" & Text8.Text & "'ORDER BY `Cardapio`.`Descricao` ASC "
                                                rs.Open sql
                                                If rs.BOF = True Then
                                                  rs.Close
                                                  sql = "SELECT `idCardapio` ,`codigo` ,`valor`FROM `Cardapio` WHERE `Descricao` LIKE '" & Text3.Text & "' AND `Medida` LIKE '" & DataGrid1.Columns(0).Value & "'AND`tipo` LIKE'" & Text8.Text & "'ORDER BY `Cardapio`.`Descricao` ASC "
                                                 rs.Open sql
                                                    If rs.BOF = True Then
                                                  rs.Close
                                                  sql = "SELECT `idCardapio` ,`codigo` ,`valor`FROM `Cardapio` WHERE `Descricao` LIKE '" & Trim(Text3.Text) & "' AND `Medida` LIKE '" & DataGrid1.Columns(0).Value & "'AND`tipo` LIKE'" & Text8.Text & "'ORDER BY `Cardapio`.`Descricao` ASC "
                                                 rs.Open sql
                                                     If rs.BOF = True Then
                                                  rs.Close
                                                  sql = "SELECT `idCardapio` ,`codigo` ,`valor`FROM `Cardapio` WHERE `Descricao` LIKE '" & Trim(Text3.Text) & "' AND `Medida` LIKE '" & Text5.Text & "' AND`tipo` LIKE'" & Text8.Text & "'ORDER BY `Cardapio`.`Descricao` ASC "
                                                 rs.Open sql
                                                   If rs.BOF = True Then
                                                  rs.Close
                                                  sql = "SELECT `idCardapio` ,`codigo` ,`valor`FROM `Cardapio` WHERE `Descricao` LIKE '" & nomeDoProduto & "' AND `Medida` LIKE '" & Text5.Text & "'AND`tipo` LIKE'" & Text8.Text & "'ORDER BY `Cardapio`.`Descricao` ASC "
                                                 rs.Open sql
                                                
                                                  If rs.BOF = True Then
                                                  rs.Close
                                                  sql = "SELECT `idCardapio` ,`codigo` ,`valor`FROM `Cardapio` WHERE `Descricao` LIKE '" & nomeDoProduto & "' AND `Medida` LIKE '" & Text5.Text & "'AND`tipo` LIKE'" & Text8.Text & "'ORDER BY `Cardapio`.`Descricao` ASC "
                                                 rs.Open sql
                                                
                                                
                                                
                                                End If
                                                
                                                End If
                                                
                                                
                                                End If
                                                
                                                
                                                End If
                                                
                                                
                                                End If
                                                
                                                
                                                
                                                
                                                If rs.BOF = False Then
                                                  Text5.Text = DataGrid1.Columns(0).Value
                                                  Text2.Text = rs.Fields("idCardapio").Value
                                                  If Not IsNull(rs.Fields("codigo").Value) Then
                                                  Text1.Text = rs.Fields("codigo").Value
                                                  End If
                                                        funcall (rs.Fields("valor").Value)
                                                                    
                                                  End If
                                                  
                                                  
                                                 'seguir ou nao para o proximo passo
                                                 If Text5.Text <> DataGrid1.Columns(0).Value Then
                                                  'contadorDePassos = contadorDePassos - 1
                                                 
                                                 
                                                 End If
                                                
                                                rs.Close
         
  
 

 'DataGrid1.Visible = F
 


 Set rs = Nothing

' Set rs = Nothing
        End Select
      
      
      
If Text5.Text = Text3.Text Then
      
reparar
End If
End Sub



Public Function funCardapioPorMeuCodigo() As Boolean
If Text5.Text = "" Then
'busca simples primerio item ou metade de uma pizza
CardapioPorMeuCodigoSemCriterio
funCardapioPorMeuCodigo = True
DataGrid1.Visible = False


Else
'buscar por metade de um determinado produto seguir os criterios de tamanho
CardapioPorMeuCodigoComCriterio
funCardapioPorMeuCodigo = False
DataGrid1.Visible = False

End If
Animation1.Close

End Function

Public Sub CardapioPorMeuCodigoSemCriterio()
If Text1.Text <> "" Then
Dim intMcodigo As Integer

intMcodigo = Trim(Text1.Text)


'ConServer



Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs.ActiveConnection = con
rs.CursorLocation = adUseClient



        Select Case Pesquisacardapiofun
            Case 0 'simples pesquisa
            
                                                 sql = "SELECT `idCardapio`,`codigo`,`Descricao` ,`tipo`,`valor`,`Medida` FROM `Cardapio` WHERE `codigo` = '" & intMcodigo & "'ORDER BY `Cardapio`.`Descricao` ASC "
                                                rs.Open sql
                                                 If rs.BOF = False Then
                                                  
                                                  Text2.Text = rs.Fields("idCardapio").Value
                                                  Text3.Text = rs.Fields("Descricao").Value
                                                     funcall (rs.Fields("valor").Value)
                                               
                                                  Text5.Text = rs.Fields("Medida").Value
                                                  'NpTitulo = rs.Fields("tipo").Value
                                                   
                                                  Else
                                                   MsgBox " O codigo : '" & intMcodigo & "' não foi localizado no sistema ", , "Não Existe"
                                                   Text1.Text = ""
                                                   Text1.SetFocus
                                                End If
                                                rs.Close
                                                   
                                                   




 Set rs = Nothing
              
            Case 1 'pesquisa por titulos
            
            
            
                                                 sql = "SELECT `idCardapio`,`codigo`,`Descricao` ,`tipo`,`valor`,`Medida` FROM `Cardapio` WHERE `codigo` = '" & intMcodigo & "'AND`Titulo` LIKE'" & Text8.Text & "'ORDER BY `Cardapio`.`Descricao` ASC "
                                                rs.Open sql
                                                 If rs.BOF = False Then
                                                  
                                                  Text2.Text = rs.Fields("idCardapio").Value
                                                  Text3.Text = rs.Fields("Descricao").Value
                                                     funcall (rs.Fields("valor").Value)
                                               
                                                  Text5.Text = rs.Fields("Medida").Value
                                                  'NpTitulo = rs.Fields("tipo").Value
                                                   
                                                  Else
                                                   MsgBox " O codigo : '" & intMcodigo & "' não foi localizado no sistemano cardápio " & Text8.Text & "", , "Não Existe"
                                                   Text1.Text = ""
                                                   Text1.SetFocus
                                                End If
                                                rs.Close
                                                   
                                                   




 Set rs = Nothing
               
            Case 2 'pesquisar subtitulos
              
                                                 sql = "SELECT `idCardapio`,`codigo`,`Descricao` ,`tipo`,`valor`,`Medida` FROM `Cardapio` WHERE `codigo` = '" & intMcodigo & "'AND`tipo` LIKE'" & Text8.Text & "'ORDER BY `Cardapio`.`Descricao` ASC "
                                                rs.Open sql
                                                 If rs.BOF = False Then
                                                  
                                                  Text2.Text = rs.Fields("idCardapio").Value
                                                  Text3.Text = rs.Fields("Descricao").Value
                                                     funcall (rs.Fields("valor").Value)
                                               
                                                  Text5.Text = rs.Fields("Medida").Value
                                                  'NpTitulo = rs.Fields("tipo").Value
                                                   
                                                  Else
                                                   MsgBox " O codigo : '" & intMcodigo & "' não foi localizado no sistema ,cardápio selecionado indice " & Text8.Text & "", , "Não Existe"
                                                   Text1.Text = ""
                                                   Text1.SetFocus
                                                End If
                                                rs.Close
                                                   
                                                   




 Set rs = Nothing
               
        End Select
Else
MsgBox "Digite o código do produto", , "Código ?"
Text1.SetFocus
End If

          
End Sub

Public Sub CardapioPorMeuCodigoComCriterio()
Dim intMcodigo As Integer
Dim Medida As String
intMcodigo = Text1.Text
Medida = Trim(Text5.Text)
'ConServer

Dim NpTitulo As String

Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs.ActiveConnection = con
rs.CursorLocation = adUseClient


 Select Case Pesquisacardapiofun
            Case 0 'simples pesquisa
            
                                                     sql = "SELECT * FROM `Cardapio` WHERE  `codigo` = '" & intMcodigo & "'"
                                                rs.Open sql
                                                 If rs.BOF = False Then
                                                               rs.Close
                                                               sql = "SELECT * FROM `Cardapio` WHERE  `codigo` = '" & intMcodigo & "' AND `Medida` LIKE '" & Medida & "'"
                                                              rs.Open sql
                                                          If rs.BOF = False Then
                                                 'sistema de compatibilidade
                                                            If Not IsNull(rs.Fields("codigo").Value) Then
                                                  Text1.Text = rs.Fields("codigo").Value
                                                  End If
                                                            
                                                            Text2.Text = rs.Fields("idCardapio").Value
                                                            Text3.Text = rs.Fields("Descricao").Value
                                                             funcall (rs.Fields("valor").Value)
                                                             
                                                            Else
                                                            MsgBox "Iten não possui compatibilidade suficiente para compor o produto", , "Iten não compatível"
                                                            Text1.Text = ""
                                                            Text1.SetFocus
                                                            End If
                                                        
                                                   Else
                                                   MsgBox " O codigo : '" & intMcodigo & "'não foi localizado no sistema ", , "Não Existe"
                                                   End If
                                                rs.Close
                                                
 Set rs = Nothing
              
            Case 1 'pesquisa por titulos
            
   
                                                     sql = "SELECT * FROM `Cardapio` WHERE  `codigo` = '" & intMcodigo & "'"
                                                rs.Open sql
                                                 If rs.BOF = False Then
                                                               rs.Close
                                                               sql = "SELECT * FROM `Cardapio` WHERE  `codigo` = '" & intMcodigo & "' AND `Medida` LIKE '" & Medida & "'AND`Titulo` LIKE'" & Text8.Text & "'"
                                                              rs.Open sql
                                                          If rs.BOF = False Then
                                                 'sistema de compatibilidade
                                                           
                                                             If Not IsNull(rs.Fields("codigo").Value) Then
                                                  Text1.Text = rs.Fields("codigo").Value
                                                  End If
                                                            Text2.Text = rs.Fields("idCardapio").Value
                                                            Text3.Text = rs.Fields("Descricao").Value
                                                             funcall (rs.Fields("valor").Value)
                                                             
                                                            Else
                                                            MsgBox "Iten não possui compatibilidade suficiente para compor o produto", , "Iten não compatível"
                                                            Text1.Text = ""
                                                            Text1.SetFocus
                                                            End If
                                                        
                                                   Else
                                                   MsgBox " O codigo : '" & intMcodigo & "'não foi localizado no sistema no cardapio " & Text8.Text & "", , "Não Existe"
                                                   End If
                                                rs.Close
                                                
 Set rs = Nothing
            
               
            Case 2 'pesquisar subtitulos
              
              
   
                                                     sql = "SELECT * FROM `Cardapio` WHERE  `codigo` = '" & intMcodigo & "'"
                                                rs.Open sql
                                                 If rs.BOF = False Then
                                                               rs.Close
                                                               sql = "SELECT * FROM `Cardapio` WHERE  `codigo` = '" & intMcodigo & "' AND `Medida` LIKE '" & Medida & "'AND`tipo` LIKE'" & Text8.Text & "'"
                                                              rs.Open sql
                                                          If rs.BOF = False Then
                                                 'sistema de compatibilidade
                                                           
                                                             If Not IsNull(rs.Fields("codigo").Value) Then
                                                  Text1.Text = rs.Fields("codigo").Value
                                                  End If
                                                            Text2.Text = rs.Fields("idCardapio").Value
                                                            Text3.Text = rs.Fields("Descricao").Value
                                                             funcall (rs.Fields("valor").Value)
                                                             
                                                            Else
                                                            MsgBox "Iten não possui compatibilidade suficiente para compor o produto", , "Iten não compatível"
                                                            Text1.Text = ""
                                                            Text1.SetFocus
                                                            End If
                                                        
                                                   Else
                                                   MsgBox " O codigo : '" & intMcodigo & "'não foi localizado no sistema , cardápio selecionado indice " & Text8.Text & "", , "Não Existe"
                                                   End If
                                                rs.Close
                                                
 Set rs = Nothing
            
                             
              
              
        End Select

                                             
End Sub



Public Function erasyVarGlobal()
 'gloValor1 As Double
 'gloValor2 As Double
' gloPassos As Integer
'gloPrimerioValorParaPrimeirMedida As String


End Function

Public Function funCardapioPorCodigoDoSistema(CODIGOsIS As Integer) As Boolean

If Text5 = "" Then
'busca comum
CardapioCodigosistem (CODIGOsIS)
funCardapioPorCodigoDoSistema = True

Else
' busque por criterios
CardapioCodigosistemComCriterio (CODIGOsIS)
funCardapioPorCodigoDoSistema = False
End If
Animation1.Close


End Function

Public Sub CardapioCodigosistemComCriterio(Mcodigo As Integer)


Dim intMcodigo As Integer
Dim valorObtido As Double
intMcodigo = Mcodigo

'ConServer

Dim NpTitulo As String

Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs.ActiveConnection = con
rs.CursorLocation = adUseClient

  Select Case Pesquisacardapiofun
            Case 0 'simples pesquisa
              
   sql = "SELECT `idCardapio`,`codigo`,`Descricao` ,`tipo`,`valor`,`Medida`FROM `Cardapio` WHERE `idCardapio` = '" & intMcodigo & "' AND `Medida` LIKE '" & Trim(Text5.Text) & "'"
                                                rs.Open sql
                                                 If rs.BOF = False Then
                                                   If Not IsNull(rs.Fields("codigo").Value) Then
                                                  Text1.Text = rs.Fields("codigo").Value
                                                  End If
                                                  Text2.Text = rs.Fields("idCardapio").Value
                                                  Text3.Text = rs.Fields("Descricao").Value
                                                  Text5.Text = rs.Fields("Medida").Value
                                                  funcall (rs.Fields("valor").Value)
                                                  
                                                                                                          
                                                End If
                                                rs.Close
         
  
  sql = "SELECT `Medida` FROM `Cardapio` WHERE `idCardapio` = '" & intMcodigo & "'ORDER BY `Cardapio`.`Descricao` ASC "
  rs.Open sql
 If rs.BOF = False Then
        If Trim(Text5.Text) <> Trim(rs.Fields("Medida").Value) Then
        
        MsgBox "O produto codigo do sistema " & Text2.Text & " não esta disponivel na medida compatível"
        End If
 
 'Set DataGrid1.DataSource = rs
 'DataGrid1.Visible = True
 'DataGrid1.SetFocus
 Else
 MsgBox " O codigo : '" & Mcodigo & "'Não foi localizado no sistema ", , "Não Existe"
 DataGrid1.Visible = False
 
 End If
 
 
 


'rs.Close
                                                 
'DataGrid1.Columns(0).Caption = NpTitulo
'DataGrid1.Columns(0).Width = 12000

 Set rs = Nothing
             
              
              
            Case 1 'pesquisa por titulos
               
               
  sql = "SELECT `idCardapio`,`codigo`,`Descricao` ,`tipo`,`valor`,`Medida`FROM `Cardapio` WHERE `idCardapio` = '" & intMcodigo & "' AND `Medida` LIKE '" & Trim(Text5.Text) & "'AND`Titulo` LIKE'" & Text8.Text & "'"
                                                rs.Open sql
                                                 If rs.BOF = False Then
                                                   If Not IsNull(rs.Fields("codigo").Value) Then
                                                  Text1.Text = rs.Fields("codigo").Value
                                                  End If
                                                  Text2.Text = rs.Fields("idCardapio").Value
                                                  Text3.Text = rs.Fields("Descricao").Value
                                                  Text5.Text = rs.Fields("Medida").Value
                                                  funcall (rs.Fields("valor").Value)
                                                  
                                                                                                          
                                                End If
                                                rs.Close
         
  
  sql = "SELECT `Medida` FROM `Cardapio` WHERE `idCardapio` = '" & intMcodigo & "'AND`Titulo` LIKE'" & Text8.Text & "'ORDER BY `Cardapio`.`Descricao` ASC "
  rs.Open sql
 If rs.BOF = False Then
        If Trim(Text5.Text) <> Trim(rs.Fields("Medida").Value) Then
        
        MsgBox "O produto codigo do sistema " & Text2.Text & " não esta disponivel na medida compatível"
        End If
 
 'Set DataGrid1.DataSource = rs
 'DataGrid1.Visible = True
 'DataGrid1.SetFocus
 Else
 MsgBox " O codigo : '" & Mcodigo & "'Não foi localizado no sistema no cardápio de " & Text8.Text & "", , "Não Existe"
 DataGrid1.Visible = False
 
 End If
 
 
 


'rs.Close
                                                 
'DataGrid1.Columns(0).Caption = NpTitulo
'DataGrid1.Columns(0).Width = 12000

 Set rs = Nothing
               
               
            Case 2 'pesquisar subtitulos
            
            
               
  sql = "SELECT `idCardapio`,`codigo`,`Descricao` ,`tipo`,`valor`,`Medida`FROM `Cardapio` WHERE `idCardapio` = '" & intMcodigo & "' AND `Medida` LIKE '" & Trim(Text5.Text) & "'AND`tipo` LIKE'" & Text8.Text & "'"
                                                rs.Open sql
                                                 If rs.BOF = False Then
                                                   If Not IsNull(rs.Fields("codigo").Value) Then
                                                  Text1.Text = rs.Fields("codigo").Value
                                                  End If
                                                  Text2.Text = rs.Fields("idCardapio").Value
                                                  Text3.Text = rs.Fields("Descricao").Value
                                                  Text5.Text = rs.Fields("Medida").Value
                                                  funcall (rs.Fields("valor").Value)
                                                  
                                                                                                          
                                                End If
                                                rs.Close
         
  
  sql = "SELECT `Medida` FROM `Cardapio` WHERE `idCardapio` = '" & intMcodigo & "'AND`tipo` LIKE'" & Text8.Text & "'ORDER BY `Cardapio`.`Descricao` ASC "
  rs.Open sql
 If rs.BOF = False Then
        If Trim(Text5.Text) <> Trim(rs.Fields("Medida").Value) Then
        
        MsgBox "O produto codigo do sistema " & Text2.Text & " não esta disponivel na medida compatível"
        End If
 
 'Set DataGrid1.DataSource = rs
 'DataGrid1.Visible = True
 'DataGrid1.SetFocus
 Else
 MsgBox " O codigo : '" & Mcodigo & "'Não foi localizado no sistema ,cardápio selecionado indice " & Text8.Text & "", , "Não Existe"
 DataGrid1.Visible = False
 
 End If
 
 
 


'rs.Close
                                                 
'DataGrid1.Columns(0).Caption = NpTitulo
'DataGrid1.Columns(0).Width = 12000

 Set rs = Nothing
            
              
        End Select


                                       
End Sub

Private Sub Text3_LostFocus()
If Label15.Visible = True Then
verificarseprecisaRecolocarFrete (Label15)
End If
End Sub
Private Sub Text1_LostFocus()
If Label15.Visible = True Then
verificarseprecisaRecolocarFrete (Label15)
End If
End Sub
Private Sub Text2_LostFocus()
If Label15.Visible = True Then
verificarseprecisaRecolocarFrete (Label15)
End If
End Sub

Private Sub Text31_Change()
If Text31.Text = 1 Then
Form23.Visible = False
Label20.Visible = True
BlocRevalide = False
Else

Label20.Visible = False
BlocRevalide = True
End If

End Sub

Private Sub Text33_Change()
If Text33.Text = "2" Then
Label22.Visible = True
Else
Label22.Visible = True
Label22.Caption = "Buscar"
End If

End Sub

Private Sub Text4_Change()
cobreOvalormaisAlto
REPARSEvalores
End Sub

Private Sub Text5_Change()

'controledeTamanhos
End Sub

Private Sub Text6_Change()
If Text6 <> "" Then
Text18.Text = Text6.Text
Set Command1.Picture = Picture2.Image
contabilize
End If
End Sub

Private Sub Text8_Change()
'Text8.Text = TEXT8.TEXT
End Sub

Private Sub Text9_Change()




If Text9.Text <> "" Then
Text17.Text = Text9.Text
'Text17 = Replace(Text9.Text, ",", ".")
' Text17 = Replace(Text17.Text, "R$", "")
' Text17 = Trim(Text17.Text)
Command5.Enabled = True
Command6.Enabled = True
contabilize
valorOriginal1 = Text9.Text
Else
Command5.Enabled = False
Command6.Enabled = False
End If
If Text9.Text <> "" Then
valorOriginal1 = Text9.Text
End If
Call Command13_Click
REPARSEvalores
End Sub

Private Sub Timer1_Timer()
Text8.Text = "Cardápio Fechado"
Pesquisacardapiofun = 0
Timer1.Interval = 0

End Sub

Private Sub Timer2_Timer()
'Text22.Text = Text22.Text & " " & Text5.Text
Timer2.Interval = 0
End Sub

Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)
SendKeys "{ENTER}"
End Sub



Private Sub TreeView1_Click()
Dim indice As Integer
Dim quantCaracterers As Integer
indice = (TreeView1.SelectedItem.Index)

  '  MsgBox (TreeView1.SelectedItem.Index)

Label11.Caption = TreeView1.Nodes.Item(indice).Key
quantCaracterers = Len(Label11.Caption)
If Text8.Text <> "" Then
        If quantCaracterers < 9 Then
        Label11.Caption = "titulo"
        Pesquisacardapiofun = 1
        treeviewTabsTitulos
        'pesquisa incluir titulo
        Else
        Label11.Caption = "subtitulos"
        treeviewTabsSubtitulos
        'pesquisas inclur subtitulos
        Pesquisacardapiofun = 2
        End If
Else
Pesquisacardapiofun = 0
End If




End Sub

Private Sub TreeView1_Collapse(ByVal Node As ComctlLib.Node)
cardapioLateral = False
'Text8.Text = "FECHADO"
Timer1.Interval = 100
Pesquisacardapiofun = 0
End Sub


Private Sub TreeView1_Expand(ByVal Node As ComctlLib.Node)
cardapioLateral = True
Text8.Text = "Cardápio aberto"
Call TreeView1_Click

End Sub

Private Sub TreeView1_Validate(Cancel As Boolean)
If cardapioLateral = True Then

Text8.Text = TreeView1.SelectedItem
End If


End Sub

Public Sub treeviewTabsSubtitulos()
Dim Proc As String

Proc = Trim("%" + TreeView1.SelectedItem + "%")
'ConServer


Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs.ActiveConnection = con
rs.CursorLocation = adUseClient


  
         



  sql = "SELECT `idCardapio`,`codigo`,`Descricao`,`Medida`,`tipo` ,`valor` FROM `Cardapio` WHERE `tipo` LIKE '" & Proc & "' ORDER BY `Cardapio`.`Descricao` ASC"
  rs.Open sql
 If rs.BOF = False Then
 Set DataGrid1.DataSource = rs
 DataGrid1.Visible = True
 DataGrid1.SetFocus
 Else
 DataGrid1.Visible = False
 
 End If



 Set rs = Nothing
'DataGrid1.Columns(0).Caption = TreeView1.SelectedItem
'DataGrid1.Columns(0).Width = 12000
End Sub




Public Sub treeviewTabsTitulos()
Dim Proc As String

Proc = Trim("%" + TreeView1.SelectedItem + "%")
'ConServer


Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs.ActiveConnection = con
rs.CursorLocation = adUseClient


  
         



  sql = "SELECT `idCardapio`,`codigo`,`Descricao`,`Medida`,`tipo` ,`valor` FROM `Cardapio` WHERE `Titulo` LIKE '" & Proc & "' ORDER BY `Cardapio`.`Descricao` ASC"
  rs.Open sql
 If rs.BOF = False Then
 Set DataGrid1.DataSource = rs
 DataGrid1.Visible = True
 DataGrid1.SetFocus
 Else
 DataGrid1.Visible = False
 
 End If



 Set rs = Nothing
'DataGrid1.Columns(0).Caption = TreeView1.SelectedItem
'DataGrid1.Columns(0).Width = 12000
End Sub




Public Sub visibleSecundarios()
Text4.Visible = True
Text6.Visible = True
Label9.Visible = True
Label10.Visible = True
End Sub
Public Sub invisibleSecundarios()
Text4.Visible = False
Text6.Visible = False
Label9.Visible = False
Label10.Visible = False
End Sub


Public Sub finalizar()

Form500.ProgressBar1.Value = 10
If Label15.Caption <> "Label15" Then
 verificarseprecisaRecolocarFrete (Label15)
 End If
Command2_Click

ultimomultiplicador = 1
Form500.ProgressBar1.Value = 50
valorOriginal2 = 0
valorOriginal1 = 0
Command5.Enabled = False
Command6.Enabled = False
Text14.Text = ""

Form500.ProgressBar1.Value = 70
invisibleSecundarios
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text6.Text = ""
Text5.Text = ""
Text9.Text = ""
Text13.Text = ""
Text15.Text = 1
'Text13.Visible = False
Text3.SetFocus
Form500.ProgressBar1.Value = 80
Text7.Visible = False
gloPassos = 1
completarFinalizar = 0
Form500.ProgressBar1.Value = 95
Text10.Text = completarFinalizar
DataGrid1.Visible = False
Form500.ProgressBar1.Value = 100
tamanHoEscolidoprimeiro = ""
MsgBox "Limpeza feita !"

End Sub


Public Sub finalizarLimparProceguir()
Form500.ProgressBar1.Value = 10
If Label15.Caption <> "Label15" Then
 verificarseprecisaRecolocarFrete (Label15)
 End If
Command2_Click

ultimomultiplicador = 1
Form500.ProgressBar1.Value = 50
valorOriginal2 = 0
valorOriginal1 = 0
Command5.Enabled = False
Command6.Enabled = False
Text14.Text = ""

Form500.ProgressBar1.Value = 70
invisibleSecundarios
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text6.Text = ""
Text5.Text = ""
Text9.Text = ""
Text13.Text = ""
Text15.Text = 1
'Text13.Visible = False
Text3.SetFocus
Form500.ProgressBar1.Value = 80
Text7.Visible = False
gloPassos = 1
completarFinalizar = 0
Form500.ProgressBar1.Value = 95
Text10.Text = completarFinalizar
DataGrid1.Visible = False
Form500.ProgressBar1.Value = 100
'MsgBox "Limpeza feita !"
Form404.Visible = True
End Sub

Public Sub reparar()
Form500.ProgressBar1.Value = 10
Command2_Click
Text15.Text = 1
ultimomultiplicador = 1
Form500.ProgressBar1.Value = 50
valorOriginal2 = 0
valorOriginal1 = 0
Command5.Enabled = False
Command6.Enabled = False
Text14.Text = ""

Form500.ProgressBar1.Value = 70
invisibleSecundarios
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text6.Text = ""
Text5.Text = ""
Text9.Text = ""
Text13.Text = ""
'Text13.Visible = False
Text3.SetFocus
Form500.ProgressBar1.Value = 80
Text7.Visible = False
gloPassos = 1
completarFinalizar = 0
Form500.ProgressBar1.Value = 95
Text10.Text = completarFinalizar
DataGrid1.Visible = False
Form500.ProgressBar1.Value = 100


End Sub


Public Sub FinalizacaoPontoFinal()
If Text9.Text <> "" Then
            'Text5.Text = "GG"
            'finalizar ou completar eis a questao
            
            ultimomultiplicador = 1
            novomultiplicador = 0

            'if
            'Text16.Text = 0
          Form500.ProgressBar1.Value = 75
          
            Text14.Text = ""
            Command5.Enabled = True
            Command6.Enabled = True
            Text13.Text = ""
            'Text13.Visible = False
            DataGrid1.Visible = False
            gloPassos = 1
            If Text7.Visible = True Then
            completarFinalizar = completarFinalizar + 0.5
            'If completarFinalizar = 1 And Text7.Text <> "" Then
            'Text7.Text = Text7.Text & "   1/2 " & Text3.Text
            'End If
            Else
            completarFinalizar = 1
            End If
            Form500.ProgressBar1.Value = 80
            If completarFinalizar = 0.5 Then
            MsgBox "completar"
          '    controledeTamanhos
              If Text15.Text <> "1" Then
              MsgBox "O produto não esta completamente montado ,escolha a quantidada ao termino do produto", , "Produto não esta completo!"
              Text15.Text = 1
              End If
            Set Command1.Picture = Picture1.Image
            visibleSecundarios
            
            gloPassos = 1
            Text1.Text = ""
            Text2.Text = ""
            Text3.Text = ""
            Text15.Text = 1
            Text3.SetFocus
            
            Else
            If Text7.Visible = True Then
            Text22.Text = Text7.Text & "  " & Text5.Text
            End If
            Form500.ProgressBar1.Value = 95
            'If Label15.Visible = False Then
            Call Command3_Click
            'End If
            Text12.Text = Format("0,00", "currency")
            Form500.ProgressBar1.Value = 95
            ''MsgBox "finalizar
            tamanHoEscolidoprimeiro = ""
            Form500.ProgressBar1.Value = 100
            invisibleSecundarios
            Text1.Text = ""
            Text2.Text = ""
            Text3.Text = ""
            Text5.Text = ""
            Text3.SetFocus
            Text9.Text = ""
            Text15.Text = 1
            
            Text7.Visible = False
            completarFinalizar = 0
            DataGrid1.Visible = False
            
            End If
            
            Text10.Text = completarFinalizar
Else
MsgBox "Você precisa continuar ate que obtenha o valor do produto ; continue com a tecla ENTER para achar o valor ", , "Valor não foi encontrado !"
If DataGrid1.Visible = True Then
DataGrid1.SetFocus
Else
If Text3.Enabled = True And Text3.Visible = True Then
Text3.SetFocus
End If

End If
End If
Form500.ProgressBar1.Value = 100
Form500.Hide
End Sub

Public Sub LimparUltimaescolha()
Dim letrasPrimeiroProduto As Integer
Dim ManteroPrimerio As String


letrasPrimeiroProduto = InStr(3, Text7.Text, " 1/2 ", 1)
'Print Len(Text7.Text)
ManteroPrimerio = Left(Text7.Text, letrasPrimeiroProduto)
   If completarFinalizar = 0.5 And Text7.Text <> "" And Text7.Visible = True And Text4.Text <> "" Then
   Text4.Text = ""
   Text6.Text = ""
If ManteroPrimerio <> "" Then
Text7.Text = Trim(ManteroPrimerio)
End If
End If
Form500.Hide
End Sub


Public Sub contabilize()
If qtdsup = "" Then
qtdsup = 1
End If
If Text18.Text <> "Text18" And Text18.Text <> "" Then
Text17 = CDbl(Text18 * qtdsup)
End If
Text14 = Format(Text18.Text, "currency")
'Dim valor1 As Double
'Dim valor2 As Double
'
'
'If Text24 <> "" Then
'
'
'
'valorAcrec = CDbl(Text16.Text)
'Else
'valorAcrec = 0
'End If
'If Text9.Text <> "" Then
'                If Text16.Text <> "" Then
'                valorAcrec = Format(Text16.Text, "General Number")
'                            If Text6.Visible = True And Text6.Text <> "" Then
'                            valor1 = Text6
'                            valor2 = Text12
'                            Text14 = Format(((valor1 + valor2 + valorAcrec) * qtdsup), "currency")
'                            Else
'                            If Text12.Text = "" Then
'                            Text12 = "0"
'                            End If
'                            valor1 = Text9
'                             If Text12.Text = "Frete aqui! pro" Then
'                            Text12.Text = 0
'                            End If
'                            valor2 = Text12
'                            If qtdsup = "" Then
'                            qtdsup = 1
'                            ElseIf qtdsup = 1 Then
'                            Text14 = Format(((valor1 + valor2 + valorAcrec) * qtdsup), "currency")
'                            Text17 = Format(((valor1 + valorAcrec) * qtdsup), "currency")
'                            Else
'                              Text14 = Format(((valor1 + valor2 + (valorAcrec * qtdsup))), "currency")
'                            Text17 = Format(((valor1 + (valorAcrec * qtdsup))), "currency")
'                            End If
'                            Text14 = Format(((valor1 + valor2 + (valorAcrec * qtdsup))), "currency")
'                            Text17 = Format(((valor1 + (valorAcrec * qtdsup))), "currency")
'                            End If
'                            Text16.Text = Format(valorAcrec, "currency")
'
'
'                Else
'
'                            If Text6.Visible = True And Text6.Text <> "" Then
'                            valor1 = Text6
'                            valor2 = Text12
'                            Text14 = Format(valor1 + valor2, "currency")
'                            Else
'                            If Text12.Text = "" Then
'                            Text12 = "0"
'                            End If
'                            valor1 = Text9
'                            If Text12.Text = "Frete aqui! pro" Then
'                            Text12.Text = 0
'                            End If
'
'                            valor2 = Text12
'                            Text14 = Format(valor1 + valor2, "currency")
'                            Text17 = Format(valor1, "currency")
'                            End If
'
'
'                End If
'Else
'MsgBox "É preciso continuar na operação até que o valor do produto esteja resolvido!", , "Tecle enter para continuar"
''SendKeys "{ENTER}"
'End If

End Sub


Public Sub verificarseprecisaRecolocarFrete(numPedido As Integer)

'

ConServer


Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs.ActiveConnection = con
rs.CursorLocation = adUseClient


  
         



  sql = "SELECT * FROM `at_frete` WHERE `fk_numPedido` = '" & numPedido & "'"
  rs.Open sql
 If rs.BOF = False Then
            If rs.Fields("frete").Value <> 0 Then
            'MsgBox "não a nessecidade de repor o frete"
            Else
           ' MsgBox "Ha necessida de reporo o frete"
            Text12.Text = LabelFRETE
            Text12.Visible = True
             End If
 Else
 MsgBox "pedido nao encontrato"
 
 End If



 Set rs = Nothing




End Sub

Public Sub CancelarPedido()
Dim numPedido As Integer
Dim NpTitulo As String
If Label15.Caption <> "Label15" Then
numPedido = Label15.Caption
'If Revalidar = True Then
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
'Revalidar = False
'ConServer

Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs.ActiveConnection = con
rs.CursorLocation = adUseClient



sql = "UPDATE `at_contadorDePedidos` SET `contador` = '0' WHERE `at_contadorDePedidos`.`id` =  '" & numPedido & "'"
                            rs.Open sql
'rs.Close
sql = "DELETE FROM `at_itens` WHERE `fk_pedido` = '" & numPedido & "' ORDER BY `id` DESC"
 rs.Open sql
 sql = "DELETE  FROM `at_Cupon` WHERE `numPedido` = '" & numPedido & "' ORDER BY `numPedido` ASC"
  rs.Open sql
sql = " DELETE FROM `Pagamento` WHERE `NUmPedido` ='" & numPedido & "' ORDER BY `NUmPedido` ASC"
 rs.Open sql
 Set rs = Nothing
End If
'comander6
'imprimirCupon
'Cancelar = True
'Form404.Text9.Text = 1
Unload Form404
Unload Form401
Unload Form402
Form2.Show
End Sub


Public Sub Revalide()
'buscar um numero que nao foi usado
If Text32.Text = 0 Then
            
            'ConServer
            'NomeDaloja = Form2.DataCombo1.Text
            
            Dim sql As String
            Dim rs As New ADODB.Recordset
            Set rs = New ADODB.Recordset
            
            Set rs.ActiveConnection = con
            
            rs.CursorLocation = adUseClient
            
            'sql = "SELECT * FROM `at_contadorDePedidos` WHERE `intPedido` LIKE '" & Form1.StatusBar1.Panels(2).Text & "' AND `contador` = 50"
            sql = "SELECT * FROM `at_contadorDePedidos` WHERE `id` = '" & Label15.Caption & "'AND `intPedido` LIKE '" & Form1.StatusBar1.Panels(2).Text & "' AND `contador` = 50"
            rs.Open sql
            If rs.BOF = False Then
            
            Else
            If BlocRevalide = True Then
                    Label15.Visible = False
                    Call Command11_Click
                    Label15.Visible = True
            End If
            
            End If
            
            Set rs = Nothing
            
            Text32.Text = 1
End If
End Sub

Public Sub limparbasedecontas()
qtdsup = 1
valorAcrec = 0
Text16.Visible = False
Text16 = 0
End Sub

Public Sub cobreOvalormaisAlto()
Dim valor1 As Double
Dim valor2 As Double
If Text4 <> "R$ 0,00" And Text4.Visible = True Then

If Text9 > Text4 Then
Text4 = Text9
Else
Text9 = Text4
End If
End If

If Text9.Text <> "" Then

valor1 = Text9

End If
If Text4.Text <> "" And Text4.Visible = True Then

valor2 = Text4

valorOriginal2 = Text4.Text
End If
 If Text16.Visible = False Then
 Text16 = 0
 End If
 If Text4.Text <> "" And Text9.Text <> "" Then
If Text9 = Text4 Then
Text6.Text = Text9 * 2
End If
 
 Text6.Text = Format(valor1 + valor2, "currency")
 'INCLUIR O ACRESSIMO
 If Text16.Text <> "" And Text16.Text <> "0" Then ' SE OVER ACRESSIMO
 Text14.Text = Format((Text6 + CDbl(Text16)), "currency")
 Text18.Text = (Text6 + CDbl(Text16))
 Text16 = Format((Text16), "currency")
 Else
 Text14.Text = Text6
 Text18.Text = valor1 + valor2
 End If
 
 End If
End Sub

Public Sub REPARSEvalores()
If Text13 <> "" Then
Obsevacao = Text13.Text
Text13.Visible = True
Text28.Text = Text13.Text & " " & Text28.Text
Else
If Text9 <> "" And Text16 <> "" Then
Text18.Text = Text9 + CDbl(Text16)
Text14 = Format((Text9 + CDbl(Text16)), "currency")
End If
End If
End Sub

Public Sub controledeTamanhos()
If Text5.Text = "" Then
tamanHoEscolidoprimeiro = ""
End If
If tamanHoEscolidoprimeiro = "" And Text5 <> "" Then
tamanHoEscolidoprimeiro = Text5.Text
End If
If Text5.Text <> "" And tamanHoEscolidoprimeiro <> "" Then
'comprare
If Text5.Text = tamanHoEscolidoprimeiro Then

Else
MsgBox "Tamanho diferente do primeiro pedaço", , "Tamanho Diferente"
tamanHoEscolidoprimeiro = ""
 Text4.Text = 0
   finalizar
End If
End If
'Timer2.Interval = 100

End Sub

Public Sub reescrevaonomedoProduto()
If Text7.Visible = False Then
'Text22.Text = Text3.Text + " " + Text5.Text
End If
End Sub
