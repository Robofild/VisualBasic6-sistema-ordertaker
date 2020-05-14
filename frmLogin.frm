VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Área Restrita aos usuário Master"
   ClientHeight    =   1995
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   6930
   Icon            =   "frmLogin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1178.712
   ScaleMode       =   0  'User
   ScaleWidth      =   6506.895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      DataField       =   "id"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   6000
      TabIndex        =   7
      Text            =   "Text2"
      Top             =   240
      Width           =   615
   End
   Begin VB.TextBox Text1 
      DataField       =   "senha"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   4680
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   600
      Width           =   1935
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   4680
      Top             =   1200
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
      Connect         =   "DSN=order_taker"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "order_taker"
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
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   495
      TabIndex        =   4
      Top             =   1020
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2100
      TabIndex        =   5
      Top             =   1020
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   525
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "&Nome de Usuário:"
      Height          =   195
      Index           =   0
      Left            =   105
      TabIndex        =   1
      Top             =   150
      Width           =   1275
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Senha:"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   540
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    'set the global var to false
    'to denote a failed login
    LoginSucceeded = False
    Dialog.Show
    
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    'check for correct password
    If txtUserName.Text <> "" And txtPassword <> "" Then
        If txtPassword = "123" Then
    
        LoginSucceeded = True
        Form7.Show
        
        Me.Hide
        Else
        MsgBox "Senha inválida, tente novamente!", , "Entrar"
        
        txtPassword.SetFocus
        ' SendKeys "{Home}+{End}"
        End If
    End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
'If KeyAscii = vbKeyReturn Then
        SendKeys ("{TAB}")
KeyAscii = 0

'If KeyAscii = 13 Then
'Call cmdOK_Click
End If


End Sub

Private Sub txtUserName_LostFocus()


ConsultarAcesso

End Sub

Public Sub ConsultarAcesso()
Dim Aula As String
Dim valor_proc As String
Dim campo As String
If txtUserName.Text <> "" Then


Adodc1.RecordSource = ""

Adodc1.CommandType = adCmdText
'SELECT * FROM `cadastro_de_usuarios_telemarketing` WHERE `nome` LIKE 'robofild'

'Adodc1.RecordSource = "SELECT * FROM Tabela_Alunos WHERE " & campo & " like '" _
                                    & valor_proc & "' ORDER BY " & campo
Adodc1.RecordSource = "SELECT * FROM `cadastro_de_usuarios_telemarketing` WHERE `nome` LIKE '" & txtUserName.Text & "'"

Adodc1.Refresh



If (Adodc1.Recordset.EOF = False) Then
End If

Else

End If


End Sub
