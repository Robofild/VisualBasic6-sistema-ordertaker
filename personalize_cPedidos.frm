VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form444 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Personalize sua tela de pedidos"
   ClientHeight    =   10155
   ClientLeft      =   4050
   ClientTop       =   2955
   ClientWidth     =   13665
   Icon            =   "personalize_cPedidos.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   10155
   ScaleWidth      =   13665
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command17 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Voltar a tela de login"
      Height          =   735
      Left            =   9480
      Picture         =   "personalize_cPedidos.frx":0A02
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   7440
      Width           =   3975
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Command16"
      Height          =   495
      Left            =   11640
      TabIndex        =   22
      Top             =   600
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "personalize_cPedidos.frx":1404
      Height          =   4575
      Left            =   15120
      TabIndex        =   21
      Top             =   2760
      Visible         =   0   'False
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   8070
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
   Begin VB.CommandButton Command15 
      Height          =   375
      Left            =   9720
      Picture         =   "personalize_cPedidos.frx":1419
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   600
      Width           =   375
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   6240
      Top             =   8640
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
      RecordSource    =   "configuracao_tela_pedido"
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
   Begin VB.CommandButton Command14 
      Caption         =   "Command14"
      Height          =   375
      Left            =   8640
      TabIndex        =   19
      Top             =   8640
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   13800
      Top             =   960
      Visible         =   0   'False
      Width           =   2280
      _ExtentX        =   4022
      _ExtentY        =   661
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
      ConnectStringType=   3
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
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   14160
      TabIndex        =   18
      Text            =   "Text2"
      Top             =   480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "personalize_cPedidos.frx":1E1B
      Height          =   360
      Left            =   4320
      TabIndex        =   17
      Top             =   600
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   635
      _Version        =   393216
      ListField       =   "nometitulo"
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10320
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox List1 
      Height          =   3375
      ItemData        =   "personalize_cPedidos.frx":1E30
      Left            =   600
      List            =   "personalize_cPedidos.frx":1E32
      TabIndex        =   15
      Top             =   4800
      Width           =   8535
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Salvar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9480
      Picture         =   "personalize_cPedidos.frx":1E34
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6480
      Width           =   3975
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H00808000&
      Caption         =   "Retirar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   11640
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3120
      Width           =   1815
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H0080FF80&
      Caption         =   "Acressimo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3120
      Width           =   1815
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00C0C0C0&
      Height          =   1335
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3120
      Width           =   1815
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0C0&
      Height          =   1335
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3120
      Width           =   1815
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0C0&
      Height          =   1335
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3120
      Width           =   1815
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0C0&
      Height          =   1335
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3120
      Width           =   1815
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0C0&
      Height          =   1335
      Left            =   11640
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0C0&
      Height          =   1335
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0C0&
      Height          =   1335
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0C0&
      Height          =   1335
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Height          =   1335
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cardapio"
      Height          =   1335
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1440
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   600
      Width           =   2295
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   375
      Left            =   6840
      Top             =   120
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
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
      RecordSource    =   "TituloCardapio"
      Caption         =   "Adodc3"
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
   Begin VB.Label Label2 
      Caption         =   "Cardápio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   16
      Top             =   360
      Width           =   1935
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   1335
      Left            =   11760
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   1335
      Left            =   840
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   1335
      Left            =   9720
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Codigo do produto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   360
      Width           =   1935
   End
End
Attribute VB_Name = "Form444"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim NumOriginalTable As Integer

Private Sub Command17_Click()
Dialog.Show

End Sub

Private Sub Command2_Click()
If DataCombo1.Text <> "" Then
Form4440.Show
Form4440.Text13.Text = NumOriginalTable
Form4440.Text2.Text = Command2.Name
'Form4440.Text13.Text = 22
'Form4440.Command1.Caption = Command2.Caption
Else
MsgBox "Antes é preciso escolher um cardápio", vbYes, "Escolha um cardapio!"
DataCombo1.SetFocus
End If

End Sub

Public Sub SelectOIndicedoCardapio(titulo As String)
Dim numIndicetitulo As Integer

ConServer

Dim sql As String
Dim rs As New ADODB.Recordset

Set rs = New ADODB.Recordset
Set rs.ActiveConnection = con

sql = "SELECT`idTituloCardapio` FROM `TituloCardapio` WHERE `nometitulo` LIKE'" & titulo & "' "

rs.Open sql

If rs.BOF = False Then
'auterado recentimente todo erro
rs.Close
sql = "SELECT `id_nomes` FROM `cli_nomes` ORDER BY `cli_nomes`.`id_nomes` ASC LIMIT 1"
rs.Open sql
NumOriginalTable = numIndicetitulo
Personalize (numIndicetitulo)

End If

 Set rs = Nothing

'protocolo do cardapio


        Form4440.Text12 = Val(numIndicetitulo)
        
        Form4440.Adodc3.RecordSource = ""
        
        Form4440.Adodc3.CommandType = adCmdText
        
        Form4440.Adodc3.RecordSource = "SELECT DISTINCT `tipo` FROM `Cardapio` WHERE `codTitulo` = '" & numIndicetitulo & "'"
        
        Form4440.Adodc3.Refresh




End Sub
Public Sub personalizetitulo(titulo As String)
Dim numIndicetitulo As Integer

ConServer

Dim sql As String
Dim rs As New ADODB.Recordset

Set rs = New ADODB.Recordset
Set rs.ActiveConnection = con

sql = "SELECT`idTituloCardapio` FROM `TituloCardapio` WHERE `nometitulo` LIKE'" & titulo & "' "

rs.Open sql

If rs.BOF = False Then
numIndicetitulo = rs.Fields("idTituloCardapio").Value

Personalize (numIndicetitulo)

End If



 Set rs = Nothing


End Sub




Private Sub Command1_Click()
Command11.Enabled = True
Command12.Enabled = True
'Form5.Show
'Form5.DataGrid1.Caption = Combo1.Text


End Sub

Private Sub Command10_Click()
If DataCombo1.Text <> "" Then
Form4440.Show
Form4440.Text13.Text = NumOriginalTable
Form4440.Text2.Text = Command10.Name
'Form4440.Text13.Text = 22
'Form4440.Command1.Caption = Command2.Caption
Else
MsgBox "Antes é preciso escolher um cardápio", vbYes, "Escolha um cardapio!"
DataCombo1.SetFocus
End If
End Sub







Private Sub Command15_Click()
Adodc3.RecordSource = ""

Adodc3.CommandType = adCmdText

Adodc3.RecordSource = "SELECT * FROM `TituloCardapio`"

Adodc3.Refresh

End Sub

Private Sub Command16_Click()
Form4440.Text12 = "1"
'Form4440.Show

End Sub



Public Sub AbilitarTodos()

Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
Command7.Enabled = True
Command8.Enabled = True
Command9.Enabled = True
Command10.Caption = ""





End Sub



Private Sub Command3_Click()
If DataCombo1.Text <> "" Then
Form4440.Show
Form4440.Text13.Text = NumOriginalTable
Form4440.Text2.Text = Command3.Name
'Form4440.Text13.Text = 22
'Form4440.Command1.Caption = Command2.Caption
Else
MsgBox "Antes é preciso escolher um cardápio", vbYes, "Escolha um cardapio!"
DataCombo1.SetFocus
End If
End Sub

Private Sub Command4_Click()
If DataCombo1.Text <> "" Then
Form4440.Show
Form4440.Text13.Text = NumOriginalTable
Form4440.Text2.Text = Command4.Name
'Form4440.Text13.Text = 22
'Form4440.Command1.Caption = Command2.Caption
Else
MsgBox "Antes é preciso escolher um cardápio", vbYes, "Escolha um cardapio!"
DataCombo1.SetFocus
End If
End Sub

Private Sub Command5_Click()
If DataCombo1.Text <> "" Then
Form4440.Show
Form4440.Text13.Text = NumOriginalTable
Form4440.Text2.Text = Command5.Name
'Form4440.Text13.Text = 22
'Form4440.Command1.Caption = Command2.Caption
Else
MsgBox "Antes é preciso escolher um cardápio", vbYes, "Escolha um cardapio!"
DataCombo1.SetFocus
End If

End Sub

Private Sub Command6_Click()
If DataCombo1.Text <> "" Then
Form4440.Show
Form4440.Text13.Text = NumOriginalTable
Form4440.Text2.Text = Command6.Name
'Form4440.Text13.Text = 22
'Form4440.Command1.Caption = Command2.Caption
Else
MsgBox "Antes é preciso escolher um cardápio", vbYes, "Escolha um cardapio!"
DataCombo1.SetFocus
End If
End Sub

Private Sub Command7_Click()
If DataCombo1.Text <> "" Then
Form4440.Show
Form4440.Text13.Text = NumOriginalTable
Form4440.Text2.Text = Command7.Name
'Form4440.Text13.Text = 22
'Form4440.Command1.Caption = Command2.Caption
Else
MsgBox "Antes é preciso escolher um cardápio", vbYes, "Escolha um cardapio!"
DataCombo1.SetFocus
End If

End Sub

Private Sub Command8_Click()
If DataCombo1.Text <> "" Then
Form4440.Show
Form4440.Text13.Text = NumOriginalTable
Form4440.Text2.Text = Command8.Name
'Form4440.Text13.Text = 22
'Form4440.Command1.Caption = Command2.Caption
Else
MsgBox "Antes é preciso escolher um cardápio", vbYes, "Escolha um cardapio!"
DataCombo1.SetFocus
End If
End Sub

Private Sub Command9_Click()
If DataCombo1.Text <> "" Then
Form4440.Show
Form4440.Text13.Text = NumOriginalTable
Form4440.Text2.Text = Command9.Name
'Form4440.Text13.Text = 22
'Form4440.Command1.Caption = Command2.Caption
Else
MsgBox "Antes é preciso escolher um cardápio", vbYes, "Escolha um cardapio!"
DataCombo1.SetFocus
End If
End Sub

Public Sub criarsemrepeticaoosTipo(titulo As String)
'Form4440.Show


Form4440.Adodc3.RecordSource = ""

Form4440.Adodc3.CommandType = adCmdText

Form4440.Adodc3.RecordSource = "SELECT DISTINCT`tipo` FROM `Cardapio` WHERE `Titulo` LIKE '" & titulo & "'"

Form4440.Adodc3.Refresh



 
End Sub

Public Sub preparecustomizeporTitulos()

Adodc3.RecordSource = ""

Adodc3.CommandType = adCmdText

Adodc3.RecordSource = "SELECT DISTINCT`Titulo` FROM `Cardapio` "

Adodc3.Refresh

 
End Sub


Public Sub Personalize(NumeroDoCardapio As Integer)
Dim numeroDeLoops As Integer
Dim nomeDoBotao  As String
Dim i As Integer


numeroDeLoops = contadordebotao(NumeroDoCardapio)

Adodc2.RecordSource = ""

Adodc2.CommandType = adCmdText

Adodc2.RecordSource = "SELECT * FROM `configuracao_tela_pedido` WHERE `CardapioOficial` = '" & NumeroDoCardapio & "'"

Adodc2.Refresh
If Adodc2.Recordset.BOF = False Then
                
                For i = 1 To numeroDeLoops Step 1
                nomeDoBotao = (Adodc2.Recordset("nomeOriginaldoBtn").Value)
                personalizaCase (nomeDoBotao)
                Adodc2.Recordset.MoveNext
                Next
                

            Else
            'nao a registro para este cardapio
            Padraodefaltbotoes

End If





End Sub

Public Sub PernsonaButton(indice As Integer)



Command2.Caption = Adodc2.Recordset("texto").Value
'fonte
Command2.FontName = Adodc2.Recordset("fonte0").Value
Command2.FontSize = Adodc2.Recordset("fonte1").Value
Command2.FontItalic = Adodc2.Recordset("fonte2").Value
Command2.FontBold = Adodc2.Recordset("fonte3").Value
Command2.FontStrikethru = Adodc2.Recordset("fonte4").Value
Command2.FontUnderline = Adodc2.Recordset("fonte5").Value
'cor
Command2.BackColor = Adodc2.Recordset("Color").Value
'imagem
Command2.Picture = LoadPicture(Adodc2.Recordset("imagemCaminho").Value)









End Sub


Public Sub customizeBotoes()
If DataCombo1.Text <> "" Then
Form4440.Show
Form4440.Text2.Text = Command2.Name
'Form4440.Text13.Text = 22
Form4440.Command1.Caption = Command2.Caption
Else
MsgBox "Antes é preciso escolher um cardápio", vbYes, "Escolha um cardapio!"
DataCombo1.SetFocus
End If
End Sub


Private Sub DataCombo1_Click(Area As Integer)
personalizetitulo (DataCombo1.Text)
End Sub

Private Sub DataCombo1_LostFocus()

SelectOIndicedoCardapio (DataCombo1.Text)
End Sub



Public Sub personalizaCase(Priorizecommand As String)
'selecione o cardapio

Select Case Priorizecommand

Case "Command2"

             If Adodc2.Recordset("texto").Value <> "" Then
            Command2.Caption = Adodc2.Recordset("texto").Value
            Else
            Command2.Caption = "Defina um nome"
            End If
            'fonte
            If Adodc2.Recordset("fonte0").Value <> "" And Adodc2.Recordset("fonte0").Value <> " " Then
            Command2.FontName = Adodc2.Recordset("fonte0").Value
            Command2.FontSize = Adodc2.Recordset("fonte1").Value
            Command2.FontItalic = Adodc2.Recordset("fonte2").Value
            Command2.FontBold = Adodc2.Recordset("fonte3").Value
            Command2.FontStrikethru = Adodc2.Recordset("fonte4").Value
            Command2.FontUnderline = Adodc2.Recordset("fonte5").Value
            End If
            'cor
            If Adodc2.Recordset("Color").Value <> "" Then
            Command2.BackColor = Adodc2.Recordset("Color").Value
            End If
            'imagem
            If Adodc2.Recordset("imagemCaminho").Value <> "" Then
            Command2.Picture = LoadPicture(Adodc2.Recordset("imagemCaminho").Value)
               Else
            Command2.Picture = LoadPicture()
            End If
            
            
Case "Command3"
             If Adodc2.Recordset("texto").Value <> "" Then
            Command3.Caption = Adodc2.Recordset("texto").Value
            Else
            Command3.Caption = "Defina um nome"
            End If
            'fonte
            If Adodc2.Recordset("fonte0").Value <> "" And Adodc2.Recordset("fonte0").Value <> " " Then
            Command3.FontName = Adodc2.Recordset("fonte0").Value
            Command3.FontSize = Adodc2.Recordset("fonte1").Value
            Command3.FontItalic = Adodc2.Recordset("fonte2").Value
            Command3.FontBold = Adodc2.Recordset("fonte3").Value
            Command3.FontStrikethru = Adodc2.Recordset("fonte4").Value
            Command3.FontUnderline = Adodc2.Recordset("fonte5").Value
            End If
            'cor
            If Adodc2.Recordset("Color").Value <> "" Then
            Command3.BackColor = Adodc2.Recordset("Color").Value
            End If
            'imagem
            If Adodc2.Recordset("imagemCaminho").Value <> "" Then
            Command3.Picture = LoadPicture(Adodc2.Recordset("imagemCaminho").Value)
               Else
            Command3.Picture = LoadPicture()
            
            End If
            
            
 Case "Command4"
             If Adodc2.Recordset("texto").Value <> "" Then
            Command4.Caption = Adodc2.Recordset("texto").Value
            Else
            Command4.Caption = "Defina um nome"
            End If
            'fonte
            If Adodc2.Recordset("fonte0").Value <> "" And Adodc2.Recordset("fonte0").Value <> " " Then
            Command4.FontName = Adodc2.Recordset("fonte0").Value
            Command4.FontSize = Adodc2.Recordset("fonte1").Value
            Command4.FontItalic = Adodc2.Recordset("fonte2").Value
            Command4.FontBold = Adodc2.Recordset("fonte3").Value
            Command4.FontStrikethru = Adodc2.Recordset("fonte4").Value
            Command4.FontUnderline = Adodc2.Recordset("fonte5").Value
            End If
            'cor
            If Adodc2.Recordset("Color").Value <> "" Then
            Command4.BackColor = Adodc2.Recordset("Color").Value
            End If
            'imagem
            If Adodc2.Recordset("imagemCaminho").Value <> "" Then
            Command4.Picture = LoadPicture(Adodc2.Recordset("imagemCaminho").Value)
                Else
            Command4.Picture = LoadPicture()
            End If


Case "Command5"
            If Adodc2.Recordset("texto").Value <> "" Then
            Command5.Caption = Adodc2.Recordset("texto").Value
            Else
            Command5.Caption = "Defina um nome"
            End If
            
            'fonte
            If Adodc2.Recordset("fonte0").Value <> "" And Adodc2.Recordset("fonte0").Value <> " " Then
            Command5.FontName = Adodc2.Recordset("fonte0").Value
            Command5.FontSize = Adodc2.Recordset("fonte1").Value
            Command5.FontItalic = Adodc2.Recordset("fonte2").Value
            Command5.FontBold = Adodc2.Recordset("fonte3").Value
            Command5.FontStrikethru = Adodc2.Recordset("fonte4").Value
            Command5.FontUnderline = Adodc2.Recordset("fonte5").Value
            End If
            'cor
            If Adodc2.Recordset("Color").Value <> "" Then
            Command5.BackColor = Adodc2.Recordset("Color").Value
            End If
            'imagem
            If Adodc2.Recordset("imagemCaminho").Value <> "" Then
            Command5.Picture = LoadPicture(Adodc2.Recordset("imagemCaminho").Value)
                 Else
            Command5.Picture = LoadPicture()
            End If

Case "Command6"
            If Adodc2.Recordset("texto").Value <> "" Then
            Command6.Caption = Adodc2.Recordset("texto").Value
            Else
            Command6.Caption = "Defina um nome"
            End If
            'fonte
            If Adodc2.Recordset("fonte0").Value <> "" And Adodc2.Recordset("fonte0").Value <> " " Then
            Command6.FontName = Adodc2.Recordset("fonte0").Value
            Command6.FontSize = Adodc2.Recordset("fonte1").Value
            Command6.FontItalic = Adodc2.Recordset("fonte2").Value
            Command6.FontBold = Adodc2.Recordset("fonte3").Value
            Command6.FontStrikethru = Adodc2.Recordset("fonte4").Value
            Command6.FontUnderline = Adodc2.Recordset("fonte5").Value
            End If
            'cor
            If Adodc2.Recordset("Color").Value <> "" Then
            Command6.BackColor = Adodc2.Recordset("Color").Value
            End If
            'imagem
            If Adodc2.Recordset("imagemCaminho").Value <> "" Then
            Command6.Picture = LoadPicture(Adodc2.Recordset("imagemCaminho").Value)
                 Else
            Command6.Picture = LoadPicture()
            End If
            
Case "Command7"
            If Adodc2.Recordset("texto").Value <> "" Then
            Command7.Caption = Adodc2.Recordset("texto").Value
            Else
            Command7.Caption = "Defina um nome"
            End If
            'fonte
            If Adodc2.Recordset("fonte0").Value <> "" And Adodc2.Recordset("fonte0").Value <> " " Then
            Command7.FontName = Adodc2.Recordset("fonte0").Value
            Command7.FontSize = Adodc2.Recordset("fonte1").Value
            Command7.FontItalic = Adodc2.Recordset("fonte2").Value
            Command7.FontBold = Adodc2.Recordset("fonte3").Value
            Command7.FontStrikethru = Adodc2.Recordset("fonte4").Value
            Command7.FontUnderline = Adodc2.Recordset("fonte5").Value
            End If
            'cor
            If Adodc2.Recordset("Color").Value <> "" Then
            Command7.BackColor = Adodc2.Recordset("Color").Value
            End If
            'imagem
            If Adodc2.Recordset("imagemCaminho").Value <> "" Then
            Command7.Picture = LoadPicture(Adodc2.Recordset("imagemCaminho").Value)
                 Else
            Command7.Picture = LoadPicture()
            End If
            
Case "Command8"
            If Adodc2.Recordset("texto").Value <> "" Then
            Command8.Caption = Adodc2.Recordset("texto").Value
            Else
            Command8.Caption = "Defina um nome"
            End If
            'fonte
            If Adodc2.Recordset("fonte0").Value <> "" And Adodc2.Recordset("fonte0").Value <> " " Then
            Command8.FontName = Adodc2.Recordset("fonte0").Value
            Command8.FontSize = Adodc2.Recordset("fonte1").Value
            Command8.FontItalic = Adodc2.Recordset("fonte2").Value
            Command8.FontBold = Adodc2.Recordset("fonte3").Value
            Command8.FontStrikethru = Adodc2.Recordset("fonte4").Value
            Command8.FontUnderline = Adodc2.Recordset("fonte5").Value
            End If
            'cor
            If Adodc2.Recordset("Color").Value <> "" Then
            Command8.BackColor = Adodc2.Recordset("Color").Value
            End If
            'imagem
            If Adodc2.Recordset("imagemCaminho").Value <> "" Then
            Command8.Picture = LoadPicture(Adodc2.Recordset("imagemCaminho").Value)
            Else
            Command8.Picture = LoadPicture()
            End If
            
Case "Command9"
            If Adodc2.Recordset("texto").Value <> "" Then
            Command9.Caption = Adodc2.Recordset("texto").Value
            Else
            Command9.Caption = "Defina um nome"
            End If
            'fonte
            If Adodc2.Recordset("fonte0").Value <> "" And Adodc2.Recordset("fonte0").Value <> " " Then
            Command9.FontName = Adodc2.Recordset("fonte0").Value
            Command9.FontSize = Adodc2.Recordset("fonte1").Value
            Command9.FontItalic = Adodc2.Recordset("fonte2").Value
            Command9.FontBold = Adodc2.Recordset("fonte3").Value
            Command9.FontStrikethru = Adodc2.Recordset("fonte4").Value
            Command9.FontUnderline = Adodc2.Recordset("fonte5").Value
            End If
            'cor
            If Adodc2.Recordset("Color").Value <> "" Then
            Command9.BackColor = Adodc2.Recordset("Color").Value
            End If
            'imagem
            If Adodc2.Recordset("imagemCaminho").Value <> "" Then
            Command9.Picture = LoadPicture(Adodc2.Recordset("imagemCaminho").Value)
                Else
            Command9.Picture = LoadPicture()
            End If
            
Case "Command10"
            If Adodc2.Recordset("texto").Value <> "" Then
            Command10.Caption = Adodc2.Recordset("texto").Value
            Else
            Command10.Caption = "Defina um nome"
            End If
            'fonte
            If Adodc2.Recordset("fonte0").Value <> "" And Adodc2.Recordset("fonte0").Value <> " " Then
            Command10.FontName = Adodc2.Recordset("fonte0").Value
            Command10.FontSize = Adodc2.Recordset("fonte1").Value
            Command10.FontItalic = Adodc2.Recordset("fonte2").Value
            Command10.FontBold = Adodc2.Recordset("fonte3").Value
            Command10.FontStrikethru = Adodc2.Recordset("fonte4").Value
            Command10.FontUnderline = Adodc2.Recordset("fonte5").Value
            End If
            'cor
            If Adodc2.Recordset("Color").Value <> "" Then
            Command10.BackColor = Adodc2.Recordset("Color").Value
            End If
            'imagem
            If Adodc2.Recordset("imagemCaminho").Value <> "" Then
            Command10.Picture = LoadPicture(Adodc2.Recordset("imagemCaminho").Value)
                 Else
            Command10.Picture = LoadPicture()
            End If

 
 
End Select









End Sub

Public Function contadordebotao(NumeroDoCardapio As Integer) As Integer



ConServer

Dim sql As String
Dim rs As New ADODB.Recordset

Set rs = New ADODB.Recordset
Set rs.ActiveConnection = con

sql = "SELECT  COUNT(*) FROM `configuracao_tela_pedido` WHERE `CardapioOficial` = '" & NumeroDoCardapio & "'"

rs.Open sql

If rs.BOF = False Then
contadordebotao = rs.Fields("COUNT(*)").Value
Else
contadordebotao = 0
End If

 Set rs = Nothing






End Function

Public Sub Padraodefaltbotoes()


            Command2.Caption = ""
            Command2.BackColor = &HC0C0C0
            
            Command3.Caption = ""
            Command3.BackColor = &HC0C0C0
            
            Command4.Caption = ""
            Command4.BackColor = &HC0C0C0
            
            
            Command5.Caption = ""
            Command5.BackColor = &HC0C0C0
            
            
            Command6.Caption = ""
            Command6.BackColor = &HC0C0C0
            
            Command7.Caption = ""
            Command7.BackColor = &HC0C0C0
            
            Command8.Caption = ""
            Command8.BackColor = &HC0C0C0
            
            Command9.Caption = ""
            Command9.BackColor = &HC0C0C0
            
            Command10.Caption = ""
            Command10.BackColor = &HC0C0C0
            
            
            





End Sub

