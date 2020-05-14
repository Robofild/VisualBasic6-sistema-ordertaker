VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form Form52 
   Caption         =   "Cardápio"
   ClientHeight    =   8610
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   16065
   LinkTopic       =   "Form9"
   ScaleHeight     =   8610
   ScaleWidth      =   16065
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   855
      Left            =   0
      TabIndex        =   4
      Text            =   "0"
      Top             =   3360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Text            =   "0"
      Top             =   2640
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   420
      Left            =   3000
      TabIndex        =   2
      Top             =   360
      Visible         =   0   'False
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   741
      _Version        =   393216
      BackColor       =   16744576
      ForeColor       =   16777215
      Text            =   "Localize o produto..."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Limpar "
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Width           =   2295
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   7215
      Left            =   480
      TabIndex        =   0
      Top             =   1080
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   12726
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   19
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
End
Attribute VB_Name = "Form52"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim titulo As String
Dim subtitulo As String
Dim produtovisualisado As Boolean

Private Sub Command1_Click()
produtovisualisado = False
titulo = ""
subtitulo = ""
DataCombo1.Visible = False
buscarTitulos
End Sub

Private Sub DataCombo1_Click(Area As Integer)
If DataCombo1.Text <> "Localize o produto..." Then
localizePorcombo (DataCombo1.Text)
End If
End Sub

Private Sub DataGrid1_Click()
If titulo = "" Then
titulo = DataGrid1.Columns(0).Value
buscarSubTitulos (DataGrid1.Columns(0).Value)
ElseIf subtitulo = "" Then
subtitulo = DataGrid1.Columns(0).Value
Else
If produtovisualisado = False Then
buscarProduto
produtovisualisado = True
End If
End If

End Sub

Private Sub DataGrid1_DblClick()
If produtovisualisado = True Then
'MsgBox DataGrid1.Columns(1).Value
If Text2.Text = 0 Then
Form50.Show
Form50.Text7.Text = DataGrid1.Columns(1).Value
Form52.Hide
Else
Form52.Hide

End If
End If
End Sub

Private Sub Form_Load()

buscarTitulos
End Sub

Public Sub buscarTitulos()
ConServer

Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs.ActiveConnection = con
rs.CursorLocation = adUseClient

'Traga os titulos

sql = "SELECT DISTINCT`Titulo` FROM `Cardapio` ORDER BY `Titulo` ASC"
rs.Open sql
Set DataGrid1.DataSource = rs


Set rs = Nothing
End Sub

Public Sub buscarSubTitulos(subtitulo As String)

Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs.ActiveConnection = con
rs.CursorLocation = adUseClient

'Traga os titulos

sql = "SELECT DISTINCT`tipo` FROM `Cardapio` WHERE `Titulo` LIKE '%" & subtitulo & "%'ORDER BY `tipo` ASC"
rs.Open sql
If rs.BOF = False Then
Set DataGrid1.DataSource = rs
Else
MsgBox "não existem cardápio para este titulo " & subtitulo & " não encontrado!", , "Não encotrado!"
titulo = ""
subtitulo = ""
buscarTitulos
End If

Set rs = Nothing

End Sub

Public Sub buscarProduto()



Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs.ActiveConnection = con
rs.CursorLocation = adUseClient

'Traga os titulos

sql = "SELECT `codigo`,`idCardapio`,`Descricao`,`Medida`,`valor` FROM `Cardapio` WHERE `Titulo` LIKE '%" & titulo & "%' AND `tipo` LIKE '%" & subtitulo & "%'ORDER BY `Descricao` ASC"
rs.Open sql
If rs.BOF = False Then
Set DataGrid1.DataSource = rs
atualizedatacombo

Else
MsgBox "não existem cardápio para este titulo " & subtitulo & " não encontrado!", , "Não encotrado!"
titulo = ""
subtitulo = ""
buscarTitulos
End If

Set rs = Nothing





End Sub

Public Sub atualizedatacombo()
DataCombo1.Visible = True
Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs.ActiveConnection = con
rs.CursorLocation = adUseClient
sql = "SELECT DISTINCT`Descricao` FROM `Cardapio` WHERE `Titulo` LIKE '%" & titulo & "%' AND `tipo` LIKE '%" & titulo & "%'ORDER BY `Descricao` ASC"
rs.Open sql
  Set DataCombo1.RowSource = rs
        
  DataCombo1.ListField = "Descricao"
End Sub


Public Sub localizePorcombo(descricao As String)


Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs.ActiveConnection = con
rs.CursorLocation = adUseClient

'Traga os titulos

sql = "SELECT `codigo`,`idCardapio`,`Descricao`,`Medida`,`valor` FROM `Cardapio` WHERE `Titulo` LIKE '%" & titulo & "%' AND `tipo` LIKE '%" & titulo & "%' AND `Descricao` LIKE '%" & descricao & "%'ORDER BY `Descricao` ASC"
rs.Open sql
If rs.BOF = False Then
Set DataGrid1.DataSource = rs


Else
MsgBox "não existem cardápio para este titulo " & subtitulo & " não encontrado!", , "Não encotrado!"
titulo = ""
subtitulo = ""
buscarTitulos
End If

Set rs = Nothing
End Sub
