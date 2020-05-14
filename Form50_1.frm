VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form50_1 
   Caption         =   "Crie um cardápio"
   ClientHeight    =   9210
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12405
   Icon            =   "Form50_1.frx":0000
   LinkTopic       =   "Form9"
   ScaleHeight     =   9210
   ScaleWidth      =   12405
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   6840
      Top             =   6600
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   1085
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
   Begin MSDataListLib.DataList DataList1 
      Height          =   4380
      Left            =   7200
      TabIndex        =   4
      Top             =   1680
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   7726
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComctlLib.TreeView TreeView1 
      Height          =   4455
      Left            =   360
      TabIndex        =   3
      Top             =   2040
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   7858
      _Version        =   327682
      Style           =   7
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Criar uma Capa"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   "Ex: Carnes ,bebidas ,vinho..."
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   1440
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "1º Passo - Escolha a capa para o cardápio "
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   3375
   End
End
Attribute VB_Name = "Form50_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nodx As Node
Dim nodr As Node
Dim capa As String



Private Sub Command1_Click()
TreeView1.LineStyle = tvwRootLines
capa = InputBox("Digite um titulo para seu cardápio", "Titulo")
Set nodx = TreeView1.Nodes.Add(, , , capa)


End Sub

Public Sub fragmentoAtrabalhar()

Dim nodx As Node
Dim nodr As Node

'Show Root Lines
TreeView1.LineStyle = tvwRootLines

'Display Checkboxes
'TreeView1.Checkboxes = True

'Add Items
Set nodx = TreeView1.Nodes.Add(, , , "capa")
nodx.Expanded = True
Set nodr = TreeView1.Nodes.Add(nodx, tvwChild, , "Sub Titulos")
nodx.Expanded = True
Set nodr = TreeView1.Nodes.Add(nodx, tvwChild, , "Inserir Itens")
Set nodx = TreeView1.Nodes.Add(, , , "Inserir itens")

nodx.Expanded = True

Set nodr = TreeView1.Nodes.Add(nodx, tvwChild, , "Item 6")
nodr.Expanded = True



Set nodx = TreeView1.Nodes.Add(, , , "Item 12")
End Sub

Private Sub TreeView1_Click()
TreeView1.LineStyle = tvwRootLines
capa = InputBox("Digite um Sub titulo para seu cardápio", "Sub Titulo")
If capa <> "" Then
nodx.Expanded = True
Set nodr = TreeView1.Nodes.Add(nodx, tvwChild, , capa)
'Set nodx = TreeView1.Nodes.Add(, , , capa)
End If
End Sub

Private Sub TreeView1_DblClick()
Debug.Print nodr
End Sub
