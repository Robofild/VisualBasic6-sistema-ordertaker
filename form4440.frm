VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form form4440 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6615
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   10275
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "form4440.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "form4440.frx":000C
   ScaleHeight     =   6615
   ScaleWidth      =   10275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1800
      Top             =   5520
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
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
      Index           =   5
      Left            =   7920
      TabIndex        =   17
      Text            =   "Text5"
      Top             =   5520
      Width           =   1575
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Index           =   4
      Left            =   7920
      TabIndex        =   16
      Text            =   "Text5"
      Top             =   4920
      Width           =   1575
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Index           =   3
      Left            =   7920
      TabIndex        =   15
      Text            =   "Text5"
      Top             =   4440
      Width           =   1575
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Index           =   2
      Left            =   7920
      TabIndex        =   14
      Text            =   "Text5"
      Top             =   3840
      Width           =   1575
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Index           =   1
      Left            =   7920
      TabIndex        =   13
      Text            =   "Text5"
      Top             =   3240
      Width           =   1575
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Index           =   0
      Left            =   7920
      TabIndex        =   12
      Text            =   "Text5"
      Top             =   2640
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   7920
      TabIndex        =   11
      Text            =   "Text4"
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   7920
      TabIndex        =   10
      Text            =   "Text3"
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   7920
      TabIndex        =   9
      Text            =   "Text2"
      Top             =   240
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Height          =   4050
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Width           =   7080
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   0
         Left            =   3720
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   960
         Width           =   2295
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Selecione..."
         Height          =   375
         Index           =   2
         Left            =   3720
         TabIndex        =   7
         Top             =   2640
         Width           =   2295
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Selecione..."
         Height          =   375
         Index           =   1
         Left            =   3720
         TabIndex        =   6
         Top             =   1800
         Width           =   2295
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   6480
         Top             =   1320
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Salvar"
         Height          =   375
         Index           =   3
         Left            =   5880
         TabIndex        =   5
         Top             =   3480
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   1935
         Left            =   600
         Picture         =   "form4440.frx":25D6
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "Texto do Botão"
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
         Left            =   3600
         TabIndex        =   4
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Fonte"
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
         Left            =   3600
         TabIndex        =   3
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Cores"
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
         Left            =   3600
         TabIndex        =   2
         Top             =   1440
         Width           =   1095
      End
   End
End
Attribute VB_Name = "form4440"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim VNomebtn As String
Dim vColor As String
Dim vText As String
Dim vFonte(5) As String

Private Sub Command2_Click(Index As Integer)
Adodc1.Recordset.AddNew
Text2.Text = VNomebtn
Text3.Text = vColor
Text4.Text = vText
Text5(0).Text = vFonte(0)
Text5(1).Text = vFonte(1)
Text5(2).Text = vFonte(2)
Text5(3).Text = vFonte(3)
Text5(4).Text = vFonte(4)
Text5(5).Text = vFonte(5)

End Sub



Private Sub Command3_Click(Index As Integer)
Command1.BackColor = colorOPT
vColor = Command1.BackColor
End Sub

Private Sub Command4_Click(Index As Integer)


form4440.CommonDialog1.ShowFont

vFonte(0) = form4440.CommonDialog1.FontName
vFonte(1) = form4440.CommonDialog1.FontSize
vFonte(2) = form4440.CommonDialog1.FontItalic
vFonte(3) = form4440.CommonDialog1.FontBold
vFonte(4) = form4440.CommonDialog1.FontStrikethru
vFonte(5) = form4440.CommonDialog1.FontUnderline
'trarar erro
On Error GoTo error

Command1.FontName = vFonte(0)
Command1.FontSize = vFonte(1)
Command1.FontItalic = vFonte(2)
Command1.FontBold = vFonte(3)
Command1.FontStrikethru = vFonte(4)
Command1.FontUnderline = vFonte(5)
Exit Sub
error:
Command1.SetFocus







End Sub

Private Sub Text1_LostFocuse(Index As Integer)
vText = Text1.Text
Command1.Caption = Text1.Text

End Sub
