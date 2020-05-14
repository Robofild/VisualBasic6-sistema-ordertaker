VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form51_2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Crie Uma cartegoria para o cardapio"
   ClientHeight    =   3690
   ClientLeft      =   8295
   ClientTop       =   4305
   ClientWidth     =   4575
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command15 
      Height          =   375
      Left            =   3720
      Picture         =   "Tipos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   240
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   3240
      Picture         =   "Tipos.frx":0A02
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   240
      Width           =   375
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   3120
      Top             =   2160
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
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
      RecordSource    =   "Cardapio_Tipo"
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
   Begin VB.CommandButton Command2 
      Caption         =   "Usar"
      Enabled         =   0   'False
      Height          =   615
      Left            =   3240
      Picture         =   "Tipos.frx":1404
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1200
      Width           =   855
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Bindings        =   "Tipos.frx":1E06
      Height          =   2820
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Insira nomes que derivarão o cardápio Ex Se o Titulo e Carnes os Tipos podem ser Churrascos , Porções  etc.."
      Top             =   240
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   4974
      _Version        =   393216
      Style           =   1
      ListField       =   "tipos"
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
   Begin VB.Label Label1 
      Caption         =   "Ex: Titulo =Carnes  então Categoria =churrascos           =Porções..."
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Width           =   4575
   End
End
Attribute VB_Name = "Form51_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim nome As String

nome = Trim(DataCombo2.Text + "%")
Adodc1.RecordSource = ""

Adodc1.CommandType = adCmdText

Adodc1.RecordSource = "SELECT * FROM `Cardapio_Tipo` WHERE `tipos` LIKE '" & nome & "'"

Adodc1.Refresh


If DataCombo2.Text <> "" Then
Command2.Enabled = True
End If
End Sub

Private Sub Command15_Click()
DataCombo2.Text = ""

Call Command1_Click
End Sub

Private Sub Command2_Click()
naoIncluirDuplicidadeeee (DataCombo2.Text)
Form51.Hide
End Sub

Private Sub DataCombo2_Click(Area As Integer)

If DataCombo2.Text <> "" Then
Command2.Enabled = True
End If


End Sub

Private Sub DataCombo2_DblClick(Area As Integer)
DataCombo2.Text = ""
End Sub

Private Sub Form_Load()
Adodc1.RecordSource = ""

Adodc1.CommandType = adCmdText
                       
Adodc1.RecordSource = "SELECT * FROM `Cardapio_Tipo` ORDER BY `Cardapio_Tipo`.`tipos` ASC"
Adodc1.Refresh

End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
'If KeyAscii = vbKeyReturn Then
       
        SendKeys ("{TAB}")
        KeyAscii = 0
        End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Form50.Text5.SetFocus
End Sub
Public Function naoIncluirDuplicidadeeee(tipo As String)



If tipo <> "" Then
tipo = Format(tipo, ">")

Adodc1.RecordSource = ""

Adodc1.CommandType = adCmdText

Adodc1.RecordSource = "SELECT * FROM `Cardapio_Tipo` WHERE `tipos` LIKE '" & tipo & "'"

Adodc1.Refresh

 


    If (Adodc1.Recordset.EOF = True) Then
    
    ultimoRegistrotitulo
    Adodc1.Recordset.AddNew
  
    Adodc1.Recordset("tipos").Value = tipo
    Adodc1.Recordset.Update
    Form50.Text10.Text = tipo
    Form50.Show
    Else
    Form50.Text11.Text = Adodc1.Recordset("idCardapio_Tipo").Value
    Form50.Text10.Text = DataCombo2.Text
    Form50.Show
    
    End If
    Form50.Text5.SetFocus
    Unload Me

Else
MsgBox "Crie ou selecione uma Categoria para seu cardápio", vbInformation, "Categoria não definida!!!"


End If




End Function

Public Function ultimoRegistrotitulo() As Integer
Adodc1.RecordSource = ""

Adodc1.CommandType = adCmdText

Adodc1.RecordSource = "SELECT * FROM `Cardapio_Tipo`"

Adodc1.Refresh
Adodc1.Recordset.MoveLast
ultimoRegistrotitulo = Adodc1.Recordset("idCardapio_Tipo").Value + 1
Form50.Text11 = ultimoRegistrotitulo

End Function

