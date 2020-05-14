VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form Dialog 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Order Taker        V17.1.0"
   ClientHeight    =   3330
   ClientLeft      =   7890
   ClientTop       =   5415
   ClientWidth     =   6045
   Icon            =   "Dialog.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   6045
   StartUpPosition =   2  'CenterScreen
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   975
      Left            =   10800
      TabIndex        =   12
      Top             =   2280
      Width           =   975
      ExtentX         =   1720
      ExtentY         =   1720
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
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
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   10800
      TabIndex        =   11
      Text            =   "Text7"
      Top             =   840
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   12000
      TabIndex        =   10
      Text            =   "Text5"
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3360
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Sair"
      Height          =   615
      Left            =   4080
      Picture         =   "Dialog.frx":0A02
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox Text4 
      DataField       =   "status"
      DataSource      =   "Adodc2"
      Height          =   375
      Left            =   7200
      TabIndex        =   8
      Text            =   "Text4"
      Top             =   2040
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      DataField       =   "data_hora"
      DataSource      =   "Adodc2"
      Height          =   285
      Left            =   7200
      TabIndex        =   7
      Text            =   "Text4"
      Top             =   1680
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      DataField       =   "id_atendente_key"
      DataSource      =   "Adodc2"
      Height          =   375
      Left            =   7200
      TabIndex        =   6
      Text            =   "Text3"
      Top             =   1200
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   7320
      Top             =   2640
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      RecordSource    =   "log_atendimento"
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
      Bindings        =   "Dialog.frx":1404
      Height          =   360
      Left            =   3120
      TabIndex        =   0
      Top             =   960
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   635
      _Version        =   393216
      BackColor       =   16777215
      ListField       =   "nome"
      Text            =   "Atendente"
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
   Begin VB.CommandButton Command1 
      Caption         =   "ok"
      Height          =   615
      Left            =   4920
      Picture         =   "Dialog.frx":1419
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3120
      MaxLength       =   6
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Delivery"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   300
      Left            =   2160
      TabIndex        =   13
      Top             =   2880
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cadastrar"
      ForeColor       =   &H80000010&
      Height          =   195
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   675
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Senha:"
      Height          =   255
      Left            =   3120
      TabIndex        =   4
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Nome:"
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   720
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   2625
      Left            =   240
      Picture         =   "Dialog.frx":1E1B
      Stretch         =   -1  'True
      Top             =   480
      Width           =   2535
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'controler
Dim situacao As Integer
Dim bloqueio As Integer
Dim taxaRepeticao As Integer
Dim aviso As Integer
Dim chave As Integer
Dim Desconectar As Integer

Dim titulo As String
Dim mensagem As String
Dim link As String
Dim ip As String
'fim controler

Dim senhaVar As String
Dim idVar As Integer



Private Sub Command1_Click()
ConServerloc


'trarar erro
On Error GoTo error




Form500.Show
Form500.Visible = False

If Text1.Text <> "" And Text1.Text = senhaVar Then
Adodc2.Recordset.AddNew
Text2.Text = idVar
Text3.Text = Now
Text4.Text = "on"
Adodc2.Recordset.Update
Adodc2.Recordset.MoveFirst
Adodc2.Refresh

'Form1.Show

Form1.Show
Form1.StatusBar1.Panels(1).Text = "Bom Trabalho! "
'Form1.StatusBar1.Panels(2).Text = DataCombo1.Text
Form1.StatusBar1.Panels.Add.Text = DataCombo1.Text


Else
    If Text1.Text <> "" Then
    MsgBox "Senha não confere ", vbCritical, "Erro de Senha"
    End If
Text1.Text = ""
Text1.SetFocus
End If
Exit Sub

error:
MsgBox "Demora no login por motivo de segurança reinicie o programa!", , "Tempo expirado!"
End
Exit Sub

End Sub

Public Sub ConsultarUsuario(nome As String)
If nome <> "" Then

ConServer
nome = Trim(nome + "%")

Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs.ActiveConnection = con
rs.CursorLocation = adUseClient
sql = "SELECT * FROM `cadastro_de_usuarios_telemarketing` WHERE `nome` LIKE '" & nome & "' ORDER BY `nome` ASC"
                                                rs.Open sql
                                                 If rs.BOF = False Then
                                                  senhaVar = rs.Fields("senha").Value
                                                         idVar = rs.Fields("id").Value
                                                 End If
                                                 
 Set rs = Nothing




 

End If



End Sub



Private Sub Command2_Click()

End
End Sub

Private Sub DataCombo1_Click(Area As Integer)
'AtualizarList
End Sub

Private Sub DataCombo1_LostFocus()
ConsultarUsuario (DataCombo1.Text)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
KeyAscii = 0
Call Command1_Click
End If

End Sub

Public Sub AtualizarList()


'Adodc1.RecordSource = ""

'Adodc1.CommandType = adCmdText

'Adodc1.RecordSource = "SELECT * FROM `cadastro_de_usuarios_telemarketing`"

'Adodc1.Refresh

 


   
End Sub





Private Sub Form_Load()
ConServer
ConServerloc
Form402.Show
Form402.Visible = False
Form403.Show
Form403.Visible = False
AtualizadaMaquina
verifiqueControler
gerarlistdeAtendente
End Sub

Private Sub Label3_Click()
Form8.Show

End Sub

Private Sub Text1_GotFocus()
   
Text1.Text = ""
End Sub

Public Sub gerarlistdeAtendente()

ConServer


Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs.ActiveConnection = con
rs.CursorLocation = adUseClient
sql = "SELECT nome FROM `cadastro_de_usuarios_telemarketing`"
                                                rs.Open sql
                                                 If rs.BOF = False Then
                                                 Set DataCombo1.RowSource = rs
                                                    'Text1.Text = rs.Fields("nome").Value
                                                 End If
                                                 
 Set rs = Nothing
End Sub

Public Sub verifiqueControler()
'Dim situacao As Integer
'Dim bloqueio As Integer
'Dim taxaRepeticao As Integer
'Dim aviso As Integer
'Dim chave As Integer
'Dim Desconectar As Integer
'
'Dim titulo As String
'Dim mensagem As String
'Dim link As String
'Dim situacao As String



ConServer


Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs.ActiveConnection = con
rs.CursorLocation = adUseClient
sql = "SELECT * FROM `controler` WHERE `chave`='1'ORDER BY `controler`.`id` DESC"
                                                rs.Open sql
                                                 If rs.BOF = False Then
                                              
                                                    situacao = rs.Fields("situacao").Value
                                                     bloqueio = rs.Fields("bloqueio").Value
                                                     taxaRepeticao = rs.Fields("taxa_repeticao").Value
                                                     aviso = rs.Fields("aviso").Value
                                                     chave = rs.Fields("chave").Value
                                                     Desconectar = rs.Fields("desconectar").Value
                                                    
                                                     titulo = rs.Fields("titulo").Value
                                                     mensagem = rs.Fields("mensagem").Value
                                                     link = rs.Fields("link").Value
                                                     situacao = rs.Fields("situacao").Value
                                                                                                      If situacao = 8 Then
                                                 rs.Close
                                                 sql = "UPDATE `controler` SET `situacao` = '0' WHERE `controler`.`id` = 2"
                                                rs.Open sql
                                                MsgBox "Sistema normalizado", , "Sucesso!"
                                               Else
                                                      crieamensagem (situacao)
                                                        End If
                                                 End If

                                                 
                                                 
 Set rs = Nothing
End Sub

Public Sub crieamensagem(optaviso As Integer)
Dim command As String
Dim winsock As String
Dim resp As Integer
Select Case optaviso

            Case 1 'simples sem manobras
            'Aviso =information
                     
            MsgBox mensagem, , titulo
            
         
              Case 2 'simples com link de pagamento direciona para o email
            'Aviso =information
            resp = MsgBox(mensagem, vbInformation, titulo)
            If resp = 1 Then
            'abrir pagina de email
            Email
            
            
            End If
            Case 3 'simples com link de pagamento direciona para link do site
            'Aviso =information
            resp = MsgBox(mensagem, vbExclamation, titulo)
            If resp = 1 Then
            'abrir pagina de email
            pagamento
            End If
            'bloqueios
            Case 4 ' pagamento passivel de bloqueio
            resp = MsgBox(mensagem, vbExclamation, titulo)
            If resp = 1 Then
            'abrir pagina de email
            pagamento
            End If
            
            Case 5 ' pagamento passivel de bloqueio nivel 2
            MsgBox mensagem, vbCritical, titulo
            resp = MsgBox(mensagem, vbCritical, titulo)
            If resp = 1 Then
            'abrir pagina de email
            pagamento
            End If
              Case 6 ' pagamento passivel de bloqueio Ativdo
            MsgBox mensagem, vbCritical, titulo
            MsgBox mensagem, vbQuestion, titulo
             resp = MsgBox(mensagem, vbCritical, titulo)
            If resp = 1 Then
            'abrir pagina de email
            pagamento
            End If
            Unload Me
            End
            Case 7 ' Erro
            MsgBox mensagem, vbCritical, titulo
            Unload Me
            End
            Case 8 ' Erro
            MsgBox mensagem, , titulo
            End
           
           Case 10 'informe ip
           
           MsgBox "Envie um arquivo de video com estas informações para robofild", , "Copie e cole "

        MsgBox "  Nome: " & Text7.Text & "  ip: " & Text5.Text & "  " & Now, vbCritical, "Fotografe este aviso ! "
        
        
            Case 11 'instar atualizador

          command = "C:\MYordertaker\at.msi"
          Shell "cmd.exe /c " & command
 
           
            Case 12 'atualiza
            verificarRegistrosdeAtualizacao
          
            
            Case 13  '
            
            
                
                
                
                
                
    End Select



End Sub
Public Sub verificarRegistrosdeAtualizacao()
Dim command As String
ip = Winsock1.LocalIP 'ip local
ConServer

Dim sql As String
Dim rs As New ADODB.Recordset
'
Set rs = New ADODB.Recordset
Set rs.ActiveConnection = con

sql = "SELECT * FROM `ATUALIZACAO` WHERE `ip` LIKE '" & ip & "'"

rs.Open sql

If rs.BOF = False Then
rs.Close
sql = "SELECT * FROM `ATUALIZACAO` WHERE `ip` LIKE '" & ip & "' AND `situacao` = 0"

rs.Open sql
            If rs.BOF = False Then
            'atualize a maquina
            
            'rs.Close
            command = "C:\MYordertaker\FTPInternetControl.exe"
            Shell "cmd.exe /c " & command
            End
            Else
            'nao Atualize a maquina
            'nao faça nada
            End If





End If


   Set rs = Nothing

End Sub
Public Sub Email()
Dim command As String
 'whatsapp
command = "C:\Order_Taker\Email.bat"
Shell "cmd.exe /c " & command
End Sub

Public Sub pagamento()
Dim command As String
 'whatsapp
command = "C:\Order_Taker\pagamento.bat"
Shell "cmd.exe /c " & command
End Sub

Public Sub AtualizadaMaquina()
Text5.Text = Winsock1.LocalIP 'ip local

Text7.Text = Winsock1.LocalHostName 'ip local


End Sub

Private Sub Text7_Change()
'WebBrowser1.Navigate Trim("http://maps.google.com/maps?q= Rua' " & retiraCaracteresEspeciais(txtRua.Text) & " ',Bairro ' " & retiraCaracteresEspeciais(tXTbAIRRO.Text) & " ',Numero ' " & retiraCaracteresEspeciais(txtnumero.Text) & " ' ,Cidade ' " & retiraCaracteresEspeciais(TXTcIDADE.Text) & " '  ")
End Sub
