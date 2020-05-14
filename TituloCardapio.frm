VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form51 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Crie uma capa para seu cardápio"
   ClientHeight    =   3720
   ClientLeft      =   5475
   ClientTop       =   4305
   ClientWidth     =   4380
   Icon            =   "TituloCardapio.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   4380
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command15 
      Height          =   375
      Left            =   3840
      Picture         =   "TituloCardapio.frx":0A02
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   240
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   3360
      Picture         =   "TituloCardapio.frx":1404
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Usar"
      Enabled         =   0   'False
      Height          =   615
      Left            =   3360
      Picture         =   "TituloCardapio.frx":1E06
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1200
      Width           =   855
   End
   Begin MSDataListLib.DataCombo DataCombo3 
      Bindings        =   "TituloCardapio.frx":2808
      Height          =   3300
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   5821
      _Version        =   393216
      Style           =   1
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
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   495
      Left            =   3240
      Top             =   3120
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
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
      Caption         =   "Titulos adc4"
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
End
Attribute VB_Name = "Form51"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim nome As String

nome = Trim(DataCombo3.Text + "%")

Adodc4.RecordSource = ""

Adodc4.CommandType = adCmdText

Adodc4.RecordSource = "SELECT * FROM `TituloCardapio` WHERE `nometitulo` LIKE '" & nome & "'"

Adodc4.Refresh

If DataCombo3.Text <> "" Then
Command2.Enabled = True


End If
End Sub

Private Sub Command15_Click()
DataCombo3.Text = ""

Call Command1_Click
End Sub

Private Sub Command2_Click()



naoIncluirDuplicidadeeee (DataCombo3.Text)
Form51.Hide
End Sub

Private Sub DataCombo3_Click(Area As Integer)

If DataCombo3.Text <> "" Then
Command2.Enabled = True
End If


End Sub

Private Sub DataCombo3_DblClick(Area As Integer)
DataCombo3.Text = ""
End Sub

Private Sub Form_Load()
Adodc4.RecordSource = ""

Adodc4.CommandType = adCmdText

                       
Adodc4.RecordSource = "SELECT * FROM `TituloCardapio` ORDER BY `TituloCardapio`.`nometitulo` ASC"
Adodc4.Refresh
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
'If KeyAscii = vbKeyReturn Then
       
        SendKeys ("{TAB}")
        KeyAscii = 0
        End If
End Sub
Public Function naoIncluirDuplicidadeeee(titulo As String)



If titulo <> "" Then
titulo = Format(titulo, ">")

Adodc4.RecordSource = ""

Adodc4.CommandType = adCmdText

Adodc4.RecordSource = "SELECT * FROM `TituloCardapio` WHERE `nometitulo` LIKE '" & titulo & "'"

Adodc4.Refresh

 


    If (Adodc4.Recordset.EOF = True) Then
    
    ultimoRegistrotitulo
    Adodc4.Recordset.AddNew
  
    Adodc4.Recordset("nometitulo").Value = titulo
    Adodc4.Recordset.Update
    Form50.Text1.Text = titulo
    Form50.Show
    Else
    Form50.Text9.Text = Adodc4.Recordset("idTituloCardapio").Value
    Form50.Text1.Text = DataCombo3.Text
    Form50.Show
    
    End If
    Form50.Text10.SetFocus
    Unload Me

Else
MsgBox "Crie ou selecione um titulo para seu cardápio", vbInformation, "Titulo não definido!!!"


End If




End Function

Public Function ultimoRegistrotitulo() As Integer
Adodc4.RecordSource = ""

Adodc4.CommandType = adCmdText

Adodc4.RecordSource = "SELECT * FROM `TituloCardapio`"

Adodc4.Refresh
Adodc4.Recordset.MoveLast
ultimoRegistrotitulo = Adodc4.Recordset("idTituloCardapio").Value + 1
Form50.Text9 = ultimoRegistrotitulo
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

'trarar erro
On Error GoTo error


Form50.Text5.SetFocus

Exit Sub

error:

Exit Sub


End Sub
