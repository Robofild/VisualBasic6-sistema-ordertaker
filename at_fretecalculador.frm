VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form23 
   BorderStyle     =   0  'None
   Caption         =   "Calculador de frete"
   ClientHeight    =   3360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6270
   Icon            =   "at_fretecalculador.frx":0000
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   6270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   6600
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   1680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   4920
      Picture         =   "at_fretecalculador.frx":0A02
      ScaleHeight     =   435
      ScaleWidth      =   795
      TabIndex        =   10
      Top             =   3600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      DragIcon        =   "at_fretecalculador.frx":1404
      Height          =   375
      Left            =   3360
      TabIndex        =   9
      Top             =   3480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Usar"
      Enabled         =   0   'False
      Height          =   615
      Left            =   4680
      Picture         =   "at_fretecalculador.frx":1E06
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2640
      Width           =   855
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Bindings        =   "at_fretecalculador.frx":2808
      Height          =   360
      Left            =   3120
      TabIndex        =   7
      Top             =   1680
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
      Bindings        =   "at_fretecalculador.frx":281D
      Height          =   360
      Left            =   360
      TabIndex        =   0
      Top             =   1680
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
      TabIndex        =   1
      Top             =   2640
      Width           =   2535
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   4200
      Top             =   3600
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
      Top             =   3600
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
      TabIndex        =   6
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "Valor"
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Bairro"
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Loja :"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Atenção você precisa ensinar o sistema quanto deve cobrar por esse frete !"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   6015
   End
End
Attribute VB_Name = "Form23"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim loja As String
Dim bairo As String
Private Sub Command1_Click()
Set Command2.Picture = Picture1.Picture
End Sub

Private Sub Command2_Click()
If Text3.Text <> "" Then

ConServer

Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset
Set rs.ActiveConnection = con
   Label1.Caption = Format(Now, "yyyy/mm/dd  hh:mm:ss ")
   'Label1.Caption = Format(Label1.Caption, "DataValue")
    sql = "INSERT INTO `robofi61_order_taker`.`re_fretebairroloja` (`bairro`, `loja`, `valor`, `user`, `created`, `modified`) VALUES ('" & Format(LTrim(DataCombo2.Text), ">") & "', '" & Format(LTrim(DataCombo1.Text), ">") & "', '" & Text3.Text & "', 'testeuse', '" & Label1.Caption & "', '" & Label1.Caption & "')"
    rs.Open sql
    
 Set rs = Nothing


Form2.Show
Form2.Label13 = Text3.Text
Unload Me

'INSERT INTO `robofi61_order_taker`.`re_fretebairroloja` (`bairro`, `loja`, `valor`, `user`, `created`, `modified`) VALUES ('bteste', 'lojatest', '15.50', 'testeuse', '00:00:00 00/00/0000', '00:00:00 00/00/0000');
'UPDATE `robofi61_order_taker`.`re_fretebairroloja` SET `bairro` = 'bteste1', `loja` = 'lojatest1', `valor` = '15.51', `user` = 'testeuse1', `modified` = '0000-00-00 00:00:01' WHERE ;
Else
MsgBox "Precisa informar o valor do frete ", , "Obrigatório valor do frete"
Form23.Visible = True
'Text3.SetFocus
'retorneaoscombos
End If
End Sub

Private Sub DataCombo1_Change()
If loja = "" Then
loja = DataCombo1.Text
End If
End Sub

Private Sub DataCombo2_Change()
If bairo = "" Then
bairo = DataCombo2.Text
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'If Text3.Text = "" Then
'MsgBox "É perciso confirmar todas as infomações antes de proseguir ", , "Complete as informações"
'retorneaoscombos
'Form23.Text3.SetFocus

'End If

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
'MsgBox "É perciso confirmar todas as infomações antes de proseguir ", , "Complete as informações"
Form23.Show
Command2.Enabled = False
'Form23.Visible = True
'Text3.SetFocus

End If



End Sub

Public Sub retorneaoscombos()
DataCombo1.Text = loja
DataCombo2.Text = bairo
End Sub
