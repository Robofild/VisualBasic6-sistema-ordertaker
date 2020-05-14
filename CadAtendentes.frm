VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form70 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de usuário"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10440
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   10440
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text5 
      DataField       =   "data"
      DataSource      =   "Adodc2"
      Height          =   375
      Left            =   7920
      TabIndex        =   32
      Text            =   "Text5"
      Top             =   4680
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text4 
      DataField       =   "user"
      DataSource      =   "Adodc2"
      Height          =   375
      Left            =   7920
      TabIndex        =   31
      Text            =   "Text4"
      Top             =   4080
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      DataField       =   "Master"
      DataSource      =   "Adodc2"
      Height          =   375
      Left            =   7920
      TabIndex        =   30
      Text            =   "Text3"
      Top             =   3480
      Visible         =   0   'False
      Width           =   1935
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   7800
      Top             =   2880
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      RecordSource    =   "master_excluindo_registros"
      Caption         =   "Exlusao"
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
   Begin VB.TextBox txt2nomeResp 
      Height          =   285
      Left            =   8040
      TabIndex        =   29
      Text            =   "Text1"
      Top             =   1440
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      DataField       =   "responsavel"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   7680
      TabIndex        =   28
      Text            =   "Text1"
      Top             =   600
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox txtValidacao 
      DataField       =   "validacao"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   9240
      TabIndex        =   21
      Top             =   1920
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cadastro"
      Height          =   3615
      Index           =   0
      Left            =   360
      TabIndex        =   15
      Top             =   2400
      Width           =   6375
      Begin VB.Frame Frame3 
         BackColor       =   &H8000000A&
         Height          =   375
         Left            =   0
         TabIndex        =   22
         Top             =   0
         Visible         =   0   'False
         Width           =   3975
         Begin VB.Label Label8 
            Caption         =   ">"
            Height          =   255
            Left            =   3000
            TabIndex        =   27
            Top             =   0
            Width           =   735
         End
         Begin VB.Label LblCout 
            AutoSize        =   -1  'True
            Caption         =   "1"
            Height          =   195
            Left            =   2400
            TabIndex        =   26
            Top             =   0
            Width           =   90
         End
         Begin VB.Label de 
            Caption         =   "Nº "
            Height          =   255
            Left            =   2040
            TabIndex        =   25
            Top             =   0
            Width           =   375
         End
         Begin VB.Label LblAtual 
            AutoSize        =   -1  'True
            Caption         =   "Label8"
            Height          =   195
            Left            =   1320
            TabIndex        =   24
            Top             =   0
            Width           =   480
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "< Nº de arquivos:"
            Height          =   195
            Left            =   0
            TabIndex        =   23
            Top             =   0
            Width           =   1230
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Atenção"
         Height          =   1815
         Left            =   120
         TabIndex        =   19
         Top             =   1440
         Width           =   3135
         Begin VB.CheckBox ChkMaster 
            Caption         =   "Cadastrar como Master"
            Height          =   255
            Left            =   360
            TabIndex        =   5
            Top             =   1200
            Width           =   2055
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Caption         =   $"CadAtendentes.frx":0000
            Height          =   1095
            Left            =   120
            TabIndex        =   20
            Top             =   360
            Width           =   2775
         End
      End
      Begin VB.TextBox txtNome 
         DataField       =   "nome"
         DataSource      =   "Adodc1"
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
         TabIndex        =   2
         Top             =   720
         Width           =   5415
      End
      Begin VB.TextBox txtSenha 
         DataField       =   "senha"
         DataSource      =   "Adodc1"
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
         IMEMode         =   3  'DISABLE
         Left            =   3480
         MaxLength       =   6
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1680
         Width           =   2295
      End
      Begin VB.TextBox TxtConfirmeSenha 
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
         IMEMode         =   3  'DISABLE
         Left            =   3480
         MaxLength       =   6
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   2400
         Width           =   2295
      End
      Begin VB.Label Label3 
         Caption         =   "Nome:"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   18
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Senha"
         Height          =   255
         Index           =   0
         Left            =   3480
         TabIndex        =   17
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Confirme a senha"
         Height          =   375
         Left            =   3480
         TabIndex        =   16
         Top             =   2160
         Width           =   1575
      End
   End
   Begin VB.TextBox txtNomeresp 
      BackColor       =   &H0080FF80&
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   1320
      Width           =   4335
   End
   Begin VB.TextBox Text2 
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
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   4800
      MaxLength       =   6
      PasswordChar    =   "*"
      TabIndex        =   11
      Text            =   "123456"
      Top             =   1320
      Width           =   2055
   End
   Begin VB.CommandButton cmdSalvar 
      Caption         =   "Salvar"
      Enabled         =   0   'False
      Height          =   615
      Left            =   4200
      Picture         =   "CadAtendentes.frx":0096
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton cmdNovo 
      Caption         =   "Novo"
      Height          =   615
      Left            =   1080
      Picture         =   "CadAtendentes.frx":0A98
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton CmdConsultar 
      Caption         =   "Consultar"
      CausesValidation=   0   'False
      Height          =   615
      Left            =   2160
      Picture         =   "CadAtendentes.frx":149A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "Editar"
      Height          =   615
      Left            =   3240
      Picture         =   "CadAtendentes.frx":1E9C
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton CmdExcluir 
      Caption         =   "Excluir"
      Enabled         =   0   'False
      Height          =   615
      Left            =   5160
      Picture         =   "CadAtendentes.frx":289E
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton CmdMoveFrist 
      Caption         =   "Mover"
      Enabled         =   0   'False
      Height          =   615
      Left            =   240
      Picture         =   "CadAtendentes.frx":32A0
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton cmdMoveLast 
      Caption         =   "Mover"
      Enabled         =   0   'False
      Height          =   615
      Left            =   6120
      Picture         =   "CadAtendentes.frx":3CA2
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   240
      Width           =   735
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   7680
      Top             =   960
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
   Begin VB.Label Label1 
      Caption         =   "Nome:"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Senha"
      Height          =   375
      Left            =   4800
      TabIndex        =   13
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   6840
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   6840
      Y1              =   960
      Y2              =   960
   End
End
Attribute VB_Name = "Form70"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim valiDacao As Integer
Dim nomeDoUserResp As String
Dim SenhaDoUserResp As String
Dim numeroDeAquivosEncontrados As Integer
Dim paginacaoDeAquivosencontrados As Integer




Private Sub ChkMaster_Click()
If ChkMaster.Value = 0 Then
    
  
    txtValidacao.Text = 0
    Else
    txtValidacao.Text = 1
    
    End If

End Sub

Private Sub CmdConsultar_Click()
enablenivel0
Dim nomeUserConsultar As String
nomeUserConsultar = InputBox("Entre com o nome do usuário", "Consultar")
ConsultarUsuario (nomeUserConsultar)
End Sub

Private Sub cmdEditar_Click()
'registar alteraçao criadas
 Alterar (1)

nomeDoUserResp = txtNomeresp.Text
cmdSalvar.Enabled = True
haBilitarparaNovoCadastro
txtNome.SetFocus

End Sub

Private Sub CmdExcluir_Click()
Excluir
End Sub

Private Sub CmdMoveFrist_Click()

If paginacaoDeAquivosencontrados >= numeroDeAquivosEncontrados Then
paginacaoDeAquivosencontrados = paginacaoDeAquivosencontrados - 1
LblCout.Caption = paginacaoDeAquivosencontrados
Adodc1.Recordset.MovePrevious
CmdMoveFrist.Enabled = True
cmdMoveLast.Enabled = True
Else
CmdMoveFrist.Enabled = False
cmdMoveLast.Enabled = True


End If
End Sub

Private Sub cmdMoveLast_Click()

If paginacaoDeAquivosencontrados < numeroDeAquivosEncontrados Then
paginacaoDeAquivosencontrados = paginacaoDeAquivosencontrados + 1
LblCout.Caption = paginacaoDeAquivosencontrados
Adodc1.Recordset.MoveNext
cmdMoveLast.Enabled = True
CmdMoveFrist.Enabled = True
Else
cmdMoveLast.Enabled = False
CmdMoveFrist.Enabled = True


End If

End Sub

Private Sub cmdNovo_Click()
    'registar alteraçao criadas
     Alterar (1)
nomeDoUserResp = txtNomeresp.Text
Adodc1.Recordset.AddNew
ChkMaster.Value = 0
txt2nomeResp = nomeDoUserResp
Text1.Text = nomeDoUserResp
Frame3.Visible = False
cmdSalvar.Enabled = True
haBilitarparaNovoCadastro
txtNome.SetFocus
End Sub

Private Sub cmdSalvar_Click()
    'registar alteraçao criadas
     Alterar (0)

 Frame3.Visible = False
    
Salvar
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
'If KeyAscii = vbKeyReturn Then
       
        SendKeys ("{TAB}")
        KeyAscii = 0
        End If
End Sub

Private Sub Form_Load()
Alterar (0)

nomeDoUserResp = txtNomeresp.Text
SenhaDoUserResp = Text2.Text
'procedimentos
If txtNomeresp <> "" Then
Consultarinicialmaster
End If
enablenivel0
End Sub
Public Sub ConsultarUsuario(nome As String)
If nome <> "" Then


Adodc1.RecordSource = ""
nome = Trim(nome + "%")
Adodc1.CommandType = adCmdText

Adodc1.RecordSource = "SELECT COUNT(*) FROM `cadastro_de_usuarios_telemarketing` WHERE `nome` LIKE '" & nome & "'"

Adodc1.Refresh
 LblAtual.Caption = Adodc1.Recordset("COUNT(*)").Value
 numeroDeAquivosEncontrados = Adodc1.Recordset("COUNT(*)").Value
 
Adodc1.RecordSource = "SELECT * FROM `cadastro_de_usuarios_telemarketing` WHERE `nome` LIKE '" & nome & "'"

Adodc1.Refresh
 


    If (Adodc1.Recordset.EOF = False) Then
    Frame3.Visible = True
    
    paginacaoDeAquivosencontrados = 1
    habilitarbotoesdemovimento
    CmdExcluir.Enabled = True

    Else
    nome = Replace(nome, "%", "")
    MsgBox "Usuário de nome : " & nome & "  não pode ser encontrado no sistema ", , "Usuário não Cadastrado!"
    paginacaoDeAquivosencontrados = 0
    numeroDeAquivosEncontrados = 0
    Frame3.Visible = False
    cmdMoveLast.Enabled = False
    CmdMoveFrist.Enabled = False
    
    
    Consultarinicialmaster
    End If



End If

End Sub
Public Sub Consultarinicialmaster()
If txtNomeresp.Text <> "" Then


Adodc1.RecordSource = ""

Adodc1.CommandType = adCmdText
Adodc1.RecordSource = "SELECT * FROM `cadastro_de_usuarios_telemarketing` WHERE `nome` LIKE '" & txtNomeresp.Text & "'"

Adodc1.Refresh
 


If (Adodc1.Recordset.EOF = False) Then
    If Adodc1.Recordset("validacao").Value = 1 Then
    ChkMaster.Value = 2
    Else
    ChkMaster.Value = 0
    End If

End If
 


End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

'perguntar se confirmar saida sem salvar

If FecharSairSalvaar(Form70.Caption) = 1 Then
Cancel = 0
UnloadMode = 0
Unload Me

Else
Cancel = 1
UnloadMode = 1
End If


'Dialog.Hide

'Dialog.Show
'Dialog.Adodc1.Recordset.Close

'Dialog.Adodc1.Refresh


End Sub

Private Sub TxtConfirmeSenha_GotFocus()
TxtConfirmeSenha.Text = ""
End Sub

Private Sub TxtConfirmeSenha_KeyPress(KeyAscii As Integer)
KeyAscii = (SoNumeros(KeyAscii))
    If KeyAscii = 0 Then
    End If
End Sub

Private Sub txtNomeresp_Change()
If txtNomeresp <> "" Then
Consultarinicialmaster
End If
End Sub

Private Sub txtsenha_Change()
If txtSenha.Enabled = False Then

TxtConfirmeSenha.Text = txtSenha.Text
End If
End Sub
Public Sub enablenivel0()
txtNomeresp.Enabled = False
txtSenha.Enabled = False
txtNome.Enabled = False
TxtConfirmeSenha.Enabled = False
ChkMaster.Enabled = False
End Sub






Public Sub haBilitarparaNovoCadastro()
txtNomeresp.Enabled = True
txtSenha.Enabled = True
txtNome.Enabled = True
TxtConfirmeSenha.Enabled = True
ChkMaster.Enabled = True
txtNomeresp.Text = nomeDoUserResp
Text2.Text = SenhaDoUserResp
End Sub

Public Sub Salvar()

nomeDoUserResp = txtNomeresp.Text
Text1.Text = nomeDoUserResp
cmdSalvar.Enabled = False
cmdNovo.Enabled = True


 If (txtSenha.Text <> "" And txtSenha.Text = TxtConfirmeSenha.Text) Then

    
'Conferir senha
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.Update
Adodc1.Refresh

Else
    If txtSenha.Text <> "" Or TxtConfirmeSenha.Text <> "" Then
    MsgBox "Senhas não confere!", vbExclamation, "Senhas não confere"
    cmdSalvar.Enabled = True
    Else
    MsgBox "Senhas em branco", vbExclamation, "Senhas vazias "
     cmdSalvar.Enabled = True
    End If







End If

cmdSalvar.Enabled = False

enablenivel0
End Sub

Private Sub txtSenha_GotFocus()
 txtSenha.Text = ""

End Sub

Private Sub txtSenha_KeyPress(KeyAscii As Integer)
KeyAscii = (SoNumeros(KeyAscii))
    If KeyAscii = 0 Then
    End If
End Sub

Private Sub txtValidacao_Change()
If txtValidacao <> "" Then
    
    If txtValidacao = 1 Then
    ChkMaster.Value = 1
    Else
    ChkMaster.Value = 0
    
    End If
End If
End Sub

Public Sub habilitarbotoesdemovimento()
If numeroDeAquivosEncontrados >= 2 Then
cmdMoveLast.Enabled = True
End If


End Sub

Public Sub paginacAo()

End Sub

Public Sub Excluir()
Dim confIrme As Integer

confIrme = MsgBox("Tem Certeza que deseja excluir o registo  " & txtNome.Text, vbCritical + vbOKCancel, "Excluir?")
 If confIrme = 1 Then
 gravarRegistroExcluidos
 Adodc1.Recordset.Delete
 Adodc1.Refresh
 End If

End Sub

Public Sub gravarRegistroExcluidos()
Adodc2.Recordset.AddNew
Text3.Text = txtNomeresp
Text4.Text = txtNome
Text5.Text = Now
Adodc2.Recordset.Update
Adodc2.Recordset.MoveFirst
Adodc2.Refresh


End Sub
