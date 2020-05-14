VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form24 
   Caption         =   "Área restrita (Cadastro e manipulação de fretes)"
   ClientHeight    =   3795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6210
   Icon            =   "at_restritoFrete.frx":0000
   LinkTopic       =   "Form9"
   Moveable        =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   6210
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1800
      TabIndex        =   10
      Top             =   3000
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Usar"
      Enabled         =   0   'False
      Height          =   615
      Left            =   4680
      Picture         =   "at_restritoFrete.frx":0A02
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      DragIcon        =   "at_restritoFrete.frx":1404
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   3840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   4920
      Picture         =   "at_restritoFrete.frx":1E06
      ScaleHeight     =   435
      ScaleWidth      =   795
      TabIndex        =   5
      Top             =   3960
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   6600
      TabIndex        =   4
      Top             =   2040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdSalvar 
      Caption         =   "Salvar"
      Enabled         =   0   'False
      Height          =   615
      Left            =   4200
      Picture         =   "at_restritoFrete.frx":2808
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton cmdNovo 
      Caption         =   "Novo"
      Height          =   615
      Left            =   1080
      Picture         =   "at_restritoFrete.frx":320A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton CmdConsultar 
      Caption         =   "Consultar"
      CausesValidation=   0   'False
      Height          =   615
      Left            =   2160
      Picture         =   "at_restritoFrete.frx":3C0C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "Editar"
      Enabled         =   0   'False
      Height          =   615
      Left            =   3240
      Picture         =   "at_restritoFrete.frx":460E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   480
      Width           =   855
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Bindings        =   "at_restritoFrete.frx":5010
      Height          =   360
      Left            =   3120
      TabIndex        =   8
      Top             =   2040
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   635
      _Version        =   393216
      ListField       =   "BAIRRO"
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
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "at_restritoFrete.frx":5025
      Height          =   360
      Left            =   360
      TabIndex        =   9
      Top             =   2040
      Width           =   2295
      _ExtentX        =   4048
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
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   4200
      Top             =   3960
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
   Begin MSAdodcLib.Adodc Adodc31 
      Height          =   330
      Left            =   360
      Top             =   3960
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
   Begin VB.Label Label1 
      Caption         =   "Atenção você precisa ensinar o sistema quanto deve cobrar por esse frete !"
      Height          =   615
      Left            =   360
      TabIndex        =   15
      Top             =   120
      Width           =   5535
   End
   Begin VB.Label Label2 
      Caption         =   "Loja :"
      Height          =   375
      Left            =   360
      TabIndex        =   14
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Bairro"
      Height          =   375
      Left            =   3120
      TabIndex        =   13
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Valor"
      Height          =   255
      Left            =   1800
      TabIndex        =   12
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "R$"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      TabIndex        =   11
      Top             =   3120
      Width           =   495
   End
End
Attribute VB_Name = "Form24"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim movimentoEditarOUNovo As Integer

Private Sub CmdConsultar_Click()
Form25.Show
cmdSalvar.Enabled = True
Form25.DataCombo1.Text = Trim(LTrim(DataCombo1.Text))
Form25.DataCombo2.Text = Trim(LTrim(DataCombo2.Text))
End Sub

Private Sub cmdEditar_Click()
movimentoEditarOUNovo = 1
abilitarbox
End Sub

Private Sub cmdNovo_Click()
ablitefuncaonovo
cmdSalvar.Enabled = True
 cmdEditar.Enabled = False
End Sub

Private Sub cmdSalvar_Click()
ConServer
Dim movimetime As String
Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset
Set rs.ActiveConnection = con


   movimetime = Format(Now, "yyyy/mm/dd  hh:mm:ss ")
   'Label1.Caption = Format(Label1.Caption, "DataValue")
   If movimentoEditarOUNovo = 0 Then
    'sql = "INSERT INTO `robofi61_order_taker`.`re_fretebairroloja` (`bairro`, `loja`, `valor`, `user`, `created`, `modified`) VALUES ('" & LTrim(DataCombo2.Text) & "', '" & LTrim(DataCombo1.Text) & "', '" & Text3.Text & "', 'testeuse', '" & movimetime & "', '" & movimetime & "')"
    sql = "INSERT INTO `robofi61_order_taker`.`re_fretebairroloja` (`bairro`, `loja`, `valor`, `user`, `created`, `modified`) VALUES ('" & Format(LTrim(DataCombo2.Text), ">") & "', '" & Format(LTrim(DataCombo1.Text), ">") & "', '" & Text3.Text & "', 'testeuse', '" & movimetime & "', '" & movimetime & "')"
    MsgBox "Frete salvo!   Bairro: " & DataCombo2.Text & " Loja:  " & DataCombo1.Text & " Valor de : " & Text3.Text & " ", , "Salvo com sucesso!"
    Else
    sql = "UPDATE  `robofi61_order_taker`.`re_fretebairroloja` SET `bairro` = '" & LTrim(DataCombo2.Text) & "',`loja` = '" & LTrim(DataCombo1.Text) & "',`valor` = '" & Text3.Text & "', `user` = 'testeuse1', `modified` = '" & movimetime & "' WHERE`id`='" & Text1.Text & "' "
    MsgBox "Frete Alterado! Bairro: " & DataCombo2.Text & " Loja:  " & DataCombo1.Text & " Valor para : " & Text3.Text & " ", , "Alterado com sucesso!"
    End If
    
    rs.Open sql
    
 Set rs = Nothing


MsgBox "Salvo loja " & DataCombo1.Text & " Para bairro " & DataCombo2, , "Salvo com sucesso!"
'cmdNovo.Enabled = False
'cmdSalvar.Enabled = False

'INSERT INTO `robofi61_order_taker`.`re_fretebairroloja` (`bairro`, `loja`, `valor`, `user`, `created`, `modified`) VALUES ('bteste', 'lojatest', '15.50', 'testeuse', '00:00:00 00/00/0000', '00:00:00 00/00/0000');
'UPDATE `UPDATE `robofi61_order_taker`.`re_fretebairroloja` SET `bairro` = 'bteste1', `loja` = 'lojatest1', `valor` = '15.51', `user` = 'testeuse1', `modified` = '0000-00-00 00:00:01' WHERE`id`='"& text1.text &"' ;

End Sub

Private Sub Command1_Click()
Set Command2.Picture = Picture1.Picture
End Sub

Private Sub Command2_Click()


Call cmdSalvar_Click
End Sub

Private Sub Form_Load()
enabliteretornodeconsulta
End Sub

Private Sub Text1_Change()
consultaporidretornando (Text1)

End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
KeyAscii = SoNumerosvirgulaPonto(KeyAscii)
End Sub

Private Sub Text3_LostFocus()
If DataCombo1.Text <> "" And DataCombo2.Text <> "" And Text3.Text <> "" Then
Text3.Text = Replace(Text3.Text, ",", ".")
Command2.Enabled = True
Command2.SetFocus
Else
MsgBox "É perciso confirmar todas as infomações antes de proseguir ", , "Complete as informações"
Command2.Enabled = False

End If


End Sub


Public Sub consultaporidretornando(id As Integer)

ConServer


Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs.ActiveConnection = con
rs.CursorLocation = adUseClient

sql = "SELECT `id`,`loja`,`bairro`,`valor`  FROM `re_fretebairroloja` WHERE `id` = '" & id & "' ORDER BY `loja` ASC"
rs.Open sql

'rs.Close

If rs.BOF = False Then
'abre a conexao
'rs.Open sql$, adocn, adOpenStatic
'define a fonte de dados para a conexao ativa
'Set DataGrid1.DataSource = rs
 DataCombo1.Text = rs.Fields("loja").Value
 DataCombo2.Text = rs.Fields("bairro").Value
 Text3.Text = rs.Fields("valor").Value
 cmdEditar.Enabled = True
 enabliteretornodeconsulta
'DataCombo1.ListField = "loja"
'DataCombo2.ListField = "bairro"
'tituloDatagrid
Else
MsgBox "Não existe fretes configurados com esses parametros, por favor tente uma nova consulta com outra configurações ", , "Não localizado"


End If
'libera a conexao
 Set rs = Nothing

End Sub

Public Sub enabliteretornodeconsulta()
DataCombo1.Enabled = False
 DataCombo2.Enabled = False
 Text3.Enabled = False

End Sub

Public Sub ablitefuncaonovo()
movimentoEditarOUNovo = 0
abilitarbox
End Sub
Public Sub abilitarbox()

DataCombo1.Enabled = True
 DataCombo2.Enabled = True
 Text3.Enabled = True
 cmdSalvar.Enabled = True
End Sub
