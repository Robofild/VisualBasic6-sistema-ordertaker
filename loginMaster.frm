VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form8 
   Caption         =   "Área Restrita aos usuário Master"
   ClientHeight    =   1500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4440
   Icon            =   "loginMaster.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form8"
   ScaleHeight     =   1500
   ScaleWidth      =   4440
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text5 
      DataField       =   "validacao"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   6240
      TabIndex        =   8
      Text            =   "Text5"
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3120
      TabIndex        =   7
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   480
      Width           =   2655
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      DataField       =   "senha"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   4575
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   450
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      DataField       =   "id"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   5895
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   90
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   4560
      Top             =   840
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
      RecordSource    =   "cadastro_de_usuarios_telemarketing"
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
   Begin VB.Label lblLabels 
      Caption         =   "&Senha:"
      Height          =   270
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   510
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "&Nome de Usuário:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1275
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
If Text5.Text = 1 Then
Form70.txtNomeresp = Text3.Text
Form70.Show


Form70.TxtConfirmeSenha = Text4.Text



Form8.Hide
Else
MsgBox "Senha inválida, tente novamente!", , "Entrar"
        Text3.Text = ""
        Text4.Text = ""
        Text3.SetFocus


End If

End Sub

Private Sub Command2_Click()
 Dialog.Show
    
    Me.Hide
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
'If KeyAscii = vbKeyReturn Then
        If Text3.Text <> "" Then
        SendKeys ("{TAB}")
        KeyAscii = 0
        ConsultarAcesso
        Else
        MsgBox "Insira um nome de usuário master", vbCritical, "Nome de usuário master não foi preechido"
        End If
        
'If KeyAscii = 13 Then
'Call cmdOK_Click
End If

End Sub

Public Sub ConsultarAcesso()

If Text3.Text <> "" Then


Adodc1.RecordSource = ""

Adodc1.CommandType = adCmdText
Adodc1.RecordSource = "SELECT * FROM `cadastro_de_usuarios_telemarketing` WHERE `nome` LIKE '" & Text3.Text & "'"

Adodc1.Refresh



If (Adodc1.Recordset.EOF = False) Then
End If

Else

End If


End Sub


Private Sub Form_Load()
Text1.Text = ""
End Sub
