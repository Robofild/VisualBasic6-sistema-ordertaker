VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form Form25 
   Caption         =   "Consultas de Frete"
   ClientHeight    =   6705
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12585
   Icon            =   "at_consultaFrete.frx":0000
   LinkTopic       =   "Form9"
   Moveable        =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   12585
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Remover"
      Height          =   615
      Left            =   9120
      Picture         =   "at_consultaFrete.frx":0A02
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Pesquisar "
      Height          =   615
      Left            =   7800
      Picture         =   "at_consultaFrete.frx":1404
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   600
      Width           =   855
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Height          =   360
      Left            =   4080
      TabIndex        =   1
      Top             =   600
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   635
      _Version        =   393216
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
      Height          =   360
      Left            =   1080
      TabIndex        =   0
      Top             =   600
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   635
      _Version        =   393216
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4815
      Left            =   960
      TabIndex        =   3
      Top             =   1440
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   8493
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
   Begin VB.Label Label2 
      Caption         =   "Bairro"
      Height          =   255
      Left            =   4200
      TabIndex        =   5
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Loja"
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "Form25"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim id As Integer
Public Sub preechergridandCombos()

ConServer


Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs.ActiveConnection = con
rs.CursorLocation = adUseClient

sql = "SELECT `id`,`loja`,`bairro`,`valor` FROM `re_fretebairroloja` ORDER BY `re_fretebairroloja`.`bairro` ASC"

rs.Open sql

'rs.Close

If rs.BOF = False Then
'abre a conexao
'rs.Open sql$, adocn, adOpenStatic
'define a fonte de dados para a conexao ativa
Set DataGrid1.DataSource = rs
Set DataCombo1.RowSource = rs
Set DataCombo2.RowSource = rs
DataCombo1.ListField = "loja"
DataCombo2.ListField = "bairro"
tituloDatagrid
End If
'libera a conexao
 Set rs = Nothing



End Sub
Public Sub consultarfretes()

ConServer


Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs.ActiveConnection = con
rs.CursorLocation = adUseClient

sql = "SELECT `id`,`loja`,`bairro`,`valor`  FROM `re_fretebairroloja` WHERE `bairro` LIKE '%" & DataCombo2.Text & "%' AND `loja` LIKE '%" & DataCombo1.Text & "%' ORDER BY `re_fretebairroloja`.`bairro` ASC"

rs.Open sql

'rs.Close

If rs.BOF = False Then
'abre a conexao
'rs.Open sql$, adocn, adOpenStatic
'define a fonte de dados para a conexao ativa
Set DataGrid1.DataSource = rs
'Set DataCombo1.RowSource = rs
'Set DataCombo2.RowSource = rs
'DataCombo1.ListField = "loja"
'DataCombo2.ListField = "bairro"
tituloDatagrid
Else
MsgBox "Não existe fretes configurados com esses parametros, por favor tente uma nova consulta com outra configurações ", , "Não localizado"


End If
'libera a conexao
 Set rs = Nothing



End Sub

Private Sub Command4_Click()
If id <> 0 Then
removerfretes
Else
MsgBox "Clique sobre o frete que pretende remover", , "Remover Frete"
DataGrid1.SetFocus
End If

End Sub

Private Sub DataGrid1_Click()
id = DataGrid1.Columns(0).Value
End Sub

Private Sub DataGrid1_DblClick()
Form24.Text1.Text = DataGrid1.Columns(0).Value

Form25.Hide

End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Form24.Text1.Text = DataGrid1.Columns(0).Value

Form25.Hide
End If
End Sub
Private Sub Command1_Click()
consultarfretes
End Sub

Private Sub Form_Load()
preechergridandCombos
End Sub

Public Sub tituloDatagrid()
DataGrid1.Columns(0).Caption = "Código"
DataGrid1.Columns(1).Caption = "Loja da encomenda"
DataGrid1.Columns(2).Caption = "Para Bairro"
DataGrid1.Columns(3).Caption = "Valor (R$)"
End Sub
Public Sub removerfretes()
'Dim  As Integer

ConServer


Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs.ActiveConnection = con
rs.CursorLocation = adUseClient

sql = "DELETE FROM `re_fretebairroloja` WHERE `id` = '" & id & "'"

rs.Open sql

Call Command1_Click
MsgBox "Frete removido com sucesso!", , "Frete removido"

 Set rs = Nothing



End Sub
