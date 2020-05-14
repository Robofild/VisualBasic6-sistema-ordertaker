VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form7 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de usuário"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13785
   Enabled         =   0   'False
   Icon            =   "CadastrodeAtendente.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   13785
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text12 
      DataField       =   "validacao"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   10920
      TabIndex        =   31
      Text            =   "Text12"
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox Text11 
      DataField       =   "responsavel"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   9960
      TabIndex        =   30
      Text            =   "Text11"
      Top             =   1800
      Width           =   1935
   End
   Begin VB.TextBox Text10 
      DataField       =   "senha"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   10200
      TabIndex        =   29
      Text            =   "Text10"
      Top             =   2880
      Width           =   1815
   End
   Begin VB.TextBox Text9 
      DataField       =   "nome"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   8160
      TabIndex        =   28
      Text            =   "Text9"
      Top             =   2280
      Width           =   3855
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   4080
      Top             =   1800
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
   Begin VB.CommandButton cmdMoveLast 
      Caption         =   "Mover"
      Enabled         =   0   'False
      Height          =   615
      Left            =   6240
      Picture         =   "CadastrodeAtendente.frx":0A02
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton CmdMoveFrist 
      Caption         =   "Mover"
      Enabled         =   0   'False
      Height          =   615
      Left            =   360
      Picture         =   "CadastrodeAtendente.frx":1404
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton CmdExcluir 
      Caption         =   "Excluir"
      Enabled         =   0   'False
      Height          =   615
      Left            =   5280
      Picture         =   "CadastrodeAtendente.frx":1E06
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "Editar"
      Enabled         =   0   'False
      Height          =   615
      Left            =   3360
      Picture         =   "CadastrodeAtendente.frx":2808
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton CmdConsultar 
      Caption         =   "Consultar"
      CausesValidation=   0   'False
      Height          =   615
      Left            =   2280
      Picture         =   "CadastrodeAtendente.frx":320A
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdNovo 
      Caption         =   "Novo"
      Height          =   615
      Left            =   1200
      Picture         =   "CadastrodeAtendente.frx":3C0C
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdSalvar 
      Caption         =   "Salvar"
      Enabled         =   0   'False
      Height          =   615
      Left            =   4320
      Picture         =   "CadastrodeAtendente.frx":460E
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   240
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cadastro de usuário master"
      Height          =   3615
      Index           =   1
      Left            =   8040
      TabIndex        =   8
      Top             =   4080
      Width           =   6015
      Begin VB.TextBox Text6 
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
         Left            =   3480
         PasswordChar    =   "*"
         TabIndex        =   23
         Text            =   "Text6"
         Top             =   2280
         Width           =   2055
      End
      Begin VB.TextBox Text5 
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
         Left            =   3480
         MaxLength       =   6
         PasswordChar    =   "*"
         TabIndex        =   22
         Text            =   "Text5"
         Top             =   1560
         Width           =   2055
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00C0FFC0&
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
         Left            =   240
         TabIndex        =   21
         Text            =   "Text4"
         Top             =   840
         Width           =   5415
      End
      Begin VB.Frame Frame2 
         Caption         =   "Atenção"
         Height          =   1815
         Left            =   240
         TabIndex        =   19
         Top             =   1440
         Width           =   3135
         Begin VB.Label Label6 
            Caption         =   $"CadastrodeAtendente.frx":5010
            Height          =   1455
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   2775
         End
      End
      Begin VB.Label Label5 
         Caption         =   "Confirme a senha"
         Height          =   375
         Left            =   3480
         TabIndex        =   11
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Nome"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   10
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Senha"
         Height          =   255
         Index           =   1
         Left            =   3480
         TabIndex        =   9
         Top             =   1320
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cadastro"
      Height          =   3615
      Index           =   0
      Left            =   600
      TabIndex        =   5
      Top             =   3240
      Width           =   6015
      Begin VB.TextBox Text8 
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
         Left            =   3360
         PasswordChar    =   "*"
         TabIndex        =   27
         Text            =   "Text8"
         Top             =   2400
         Width           =   2295
      End
      Begin VB.TextBox Text7 
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
         Left            =   3360
         MaxLength       =   6
         PasswordChar    =   "*"
         TabIndex        =   25
         Text            =   "Text7"
         Top             =   1680
         Width           =   2295
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
         Height          =   375
         Left            =   240
         TabIndex        =   24
         Text            =   "Text3"
         Top             =   840
         Width           =   5415
      End
      Begin VB.Label Label7 
         Caption         =   "Confirme a senha"
         Height          =   375
         Left            =   3360
         TabIndex        =   26
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Senha"
         Height          =   255
         Index           =   0
         Left            =   3360
         TabIndex        =   7
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Nome:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   600
         Width           =   1335
      End
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   4815
      Left            =   360
      TabIndex        =   4
      Top             =   2520
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   8493
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Usuario"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Master"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
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
      Left            =   4920
      MaxLength       =   6
      PasswordChar    =   "*"
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   1320
      Width           =   2055
   End
   Begin VB.TextBox Text1 
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
      Left            =   240
      TabIndex        =   1
      Text            =   "robofild"
      Top             =   1320
      Width           =   4335
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   240
      X2              =   6960
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   6960
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label Label2 
      Caption         =   "Senha"
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Nome:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim valiDacao As Integer
Dim nomeDoUserResp As String
Dim SenhaDoUserResp As Integer




Private Sub Command1_Click(index As Integer)
If (index = 0) Then

MsgBox "ok"
Else

MsgBox "ok2"
End If
End Sub

Private Sub CmdConsultar_Click()
Dim nomeUserConsultar As String
nomeUserConsultar = InputBox("Entre com o nome do usuário", "Consultar")
ConsultarUsuario (nomeUserConsultar)

End Sub

Private Sub CmdMoveFrist_Click()
If Adodc1.Recordset.BOF = False Then


Adodc1.Recordset.MovePrevious
Else
CmdMoveFrist.Enabled = False


End If
cmdMoveLast.Enabled = True

End Sub

Private Sub cmdMoveLast_Click()
If Adodc1.Recordset.BOF = False Then

Adodc1.Recordset.MoveNext

Else
cmdMoveLast.Enabled = False


End If
CmdMoveFrist.Enabled = True

End Sub

Private Sub cmdNovo_Click()
Adodc1.Recordset.AddNew
cmdNovo.Enabled = False
AbilitarboxparaAlteracao (0)
AbilitarboxparaAlteracao (1)
direcionarParaAba1
cmdSalvar.Enabled = True



End Sub

Private Sub cmdSalvar_Click()
cmdSalvar.Enabled = False
cmdNovo.Enabled = True

 If (valiDacao = 0 And Text8.Text = Text7.Text And Text7.Text <> "") Or (valiDacao = 1 And Text5.Text = Text6.Text And Text5.Text <> "") Then

    If valiDacao = 1 Then
    Text11.Text = Text1.Text
    Text9.Text = Text4.Text
    Text10.Text = Text5.Text
    Text12.Text = valiDacao

    Else

    Text11.Text = Text1.Text
    Text9.Text = Text3.Text
    Text10.Text = Text7.Text
    Text12.Text = valiDacao
    End If

'Conferir senha
Adodc1.Recordset.Update
Adodc1.Recordset.MoveFirst

Else
    If Text7.Text <> "" Or Text8.Text <> "" Then
    MsgBox "Senhas não confere!", vbExclamation, "Senhas não confere"
    cmdSalvar.Enabled = True
    Else
    MsgBox "Senhas em branco", vbExclamation, "Senhas vazias "
     cmdSalvar.Enabled = True
    End If
If valiDacao = 1 Then
Text5.Text = ""
Text6.Text = ""
Text5.SetFocus
Else

Text7.Text = ""
Text8.Text = ""
Text7.SetFocus

End If






End If


End Sub

Private Sub Form_Load()
nomeDoUserResp = Text1.Text
limparformulario

Consultarinicialmaster
enablenivel0
 Frame1(0).ZOrder
 ajusta_container
End Sub

Private Sub TabStrip1_Click()
Dim i As Integer

i = TabStrip1.SelectedItem.index

Frame1(i - 1).ZOrder

    If (i = 2) Then
    Frame1(1).Visible = True
    Frame1(0).Visible = False
    valiDacao = 1
    If Text4.Enabled = True Then
    Text4.SetFocus
    End If
    Else
    
    valiDacao = 0
    Frame1(0).Visible = True
    Frame1(1).Visible = False
    
    If Text3.Enabled = True Then
    Text3.SetFocus
    End If
    
    End If

End Sub
Private Sub ajusta_container()
Dim i As Integer
With TabStrip1
For i = 1 To .Tabs.Count
Frame1(i - 1).Move .ClientLeft, .ClientTop, .ClientWidth, .ClientHeight
Next
End With
End Sub

Public Sub enablenivel0()
Text3.Enabled = False


Text7.Enabled = False
Text8.Enabled = False

Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False





End Sub


Public Sub limparformulario()
Text3.Text = ""
Text7.Text = ""
Text8.Text = ""

Text4.Text = ""
Text5.Text = ""
Text6.Text = ""





End Sub

Public Sub Consultarinicialmaster()
If Text1.Text <> "" Then


Adodc1.RecordSource = ""

Adodc1.CommandType = adCmdText
Adodc1.RecordSource = "SELECT * FROM `cadastro_de_usuarios_telemarketing` WHERE `nome` LIKE '" & Text1.Text & "'"

Adodc1.Refresh
 


If (Adodc1.Recordset.EOF = False) Then
    If Adodc1.Recordset("validacao").Value = 1 Then
    valiDacao = 1
       
    TabStrip1.Tabs(2).Selected = True
    Frame1(1).Visible = True
    Frame1(0).Visible = False
    Else
     TabStrip1.Tabs(1).Selected = True
    valiDacao = 0
       
    Frame1(0).Visible = True
    Frame1(1).Visible = False
    
    




     
    
    End If
End If

Else

End If
PassarvaloresAsConsulta
End Sub

Private Sub Text5_Change()
Text6.Text = Text5.Text

End Sub

Public Sub ConsultarUsuario(nome As String)
If nome <> "" Then


Adodc1.RecordSource = ""
nome = Trim(nome + "%")
Adodc1.CommandType = adCmdText
Adodc1.RecordSource = "SELECT * FROM `cadastro_de_usuarios_telemarketing` WHERE `nome` LIKE '" & nome & "'"

Adodc1.Refresh
 


If (Adodc1.Recordset.EOF = False) Then
AbilitarbotoesReferentesConsulta
' caso retorn on
      If Adodc1.Recordset("validacao").Value = 1 Then
      AbilitarboxparaAlteracao (1)
      Else
      AbilitarboxparaAlteracao (0)
      End If


    If Adodc1.Recordset("validacao").Value = 1 Then
       valiDacao = 1
       
    TabStrip1.Tabs(2).Selected = True
    Frame1(1).Visible = True
    Frame1(0).Visible = False
    Else
    TabStrip1.Tabs(1).Selected = True
     valiDacao = 0
       
    Frame1(0).Visible = True
    Frame1(1).Visible = False
    End If
Else
   nome = Replace(nome, "%", "")
   MsgBox "Usuário de nome : " & nome & "  não pode ser encontrado no sistema ", , "Usuário não Cadastrado!"
   Consultarinicialmaster
End If

Else

End If
PassarvaloresAsConsulta
End Sub

Public Sub AbilitarboxparaAlteracao(tipo As Integer)
    If (tipo = 1) Then
    TabStrip1.Tabs(2).Selected = True
    Frame1(1).Visible = True
    Frame1(0).Visible = False
    Else
     TabStrip1.Tabs(1).Selected = True
     valiDacao = 0
       
    Frame1(0).Visible = True
    Frame1(1).Visible = False
    End If
If (tipo = 1) Then
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text4.SetFocus

Else
Text3.Enabled = True
Text7.Enabled = True
Text8.Enabled = True
Text3.SetFocus

End If


End Sub

Public Sub AbilitarbotoesReferentesConsulta()
cmdMoveLast.Enabled = True
cmdEditar.Enabled = True
cmdSalvar.Enabled = True
CmdExcluir.Enabled = True
CmdMoveFrist.Enabled = True




End Sub


Public Sub PassarvaloresAsConsulta()
If valiDacao = 1 Then
Text4.Text = Text9.Text
Text5.Text = Text10.Text
Else
Text3.Text = Text9.Text
Text10.Text = Text7.Text

End If
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)


KeyAscii = (SoNumeros(KeyAscii))

If KeyAscii = 0 Then



End If

End Sub

Private Sub Text9_Change()
PassarvaloresAsConsulta
End Sub

Public Sub direcionarParaAba1()
TabStrip1.Tabs(1).Selected = True
    valiDacao = 0
       
    Frame1(0).Visible = True
    Frame1(1).Visible = False
End Sub
