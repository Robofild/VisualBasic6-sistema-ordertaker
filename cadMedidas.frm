VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form51_3 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Crie medida para esse Produto"
   ClientHeight    =   3705
   ClientLeft      =   9435
   ClientTop       =   7590
   ClientWidth     =   4020
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   4020
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command15 
      Height          =   375
      Left            =   3480
      Picture         =   "cadMedidas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   240
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   3000
      Picture         =   "cadMedidas.frx":0A02
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   240
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Usar"
      Enabled         =   0   'False
      Height          =   615
      Left            =   3000
      Picture         =   "cadMedidas.frx":1404
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1200
      Width           =   855
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "cadMedidas.frx":1E06
      Height          =   2820
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Você pode e deve criar a medida que mais atende ao produto em questão!"
      Top             =   240
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   4974
      _Version        =   393216
      Style           =   1
      ListField       =   "unidade"
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
      Height          =   495
      Left            =   2760
      Top             =   3000
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      RecordSource    =   "Cardapio_medidas"
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
      Caption         =   "Ex:Grande,pequena, Kilos...."
      Height          =   1095
      Left            =   240
      TabIndex        =   2
      Top             =   3120
      Width           =   1575
   End
End
Attribute VB_Name = "Form51_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim nome As String

nome = Trim(DataCombo1.Text + "%")
Adodc1.RecordSource = ""

Adodc1.CommandType = adCmdText
                       
Adodc1.RecordSource = "SELECT * FROM `Cardapio_medidas` WHERE `unidade` LIKE '" & nome & "'"
Adodc1.Refresh

 
If DataCombo1.Text <> "" Then
Command2.Enabled = True
End If
End Sub

Private Sub Command15_Click()
DataCombo1.Text = ""

Call Command1_Click
End Sub

Private Sub Command2_Click()
naoIncluirDuplicidadeeee (DataCombo1.Text)
Form51.Hide
End Sub

Private Sub DataCombo1_Click(Area As Integer)

If DataCombo1.Text <> "" Then
Command2.Enabled = True
End If


End Sub

Private Sub DataCombo1_DblClick(Area As Integer)
DataCombo1.Text = ""
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
'If KeyAscii = vbKeyReturn Then
       
        SendKeys ("{TAB}")
        KeyAscii = 0
        End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)


'trarar erro
On Error GoTo error
Form50.Text5.SetFocus
Exit Sub
error:

Exit Sub
End Sub

Private Sub Form_Load()
Adodc1.RecordSource = ""

Adodc1.CommandType = adCmdText
                       
Adodc1.RecordSource = "SELECT * FROM `Cardapio_medidas` ORDER BY `Cardapio_medidas`.`unidade` ASC"
Adodc1.Refresh
End Sub

Public Function naoIncluirDuplicidadeeee(Medida As String)



If Medida <> "" Then
Medida = Format(Medida, ">")

Adodc1.RecordSource = ""

Adodc1.CommandType = adCmdText
                       
Adodc1.RecordSource = "SELECT * FROM `Cardapio_medidas` WHERE `unidade` LIKE '" & Medida & "'"
Adodc1.Refresh

 


    If (Adodc1.Recordset.EOF = True) Then
    
    ultimoRegistrotitulo
    Adodc1.Recordset.AddNew
  
    Adodc1.Recordset("unidade").Value = Medida
    Adodc1.Recordset.Update
    Form50.Text10.Text = Medida
    Form50.Show
    Else
    Form50.Text13.Text = Adodc1.Recordset("idCardapio_medidas").Value
    Form50.Text12.Text = DataCombo1.Text
    Form50.Show
    
    End If
    Form50.Text4.SetFocus
    Form51.Hide
    

Else
MsgBox "Crie ou selecione uma medida para seu produto", vbInformation, "Medida não definida!!!"


End If




End Function

Public Function ultimoRegistrotitulo() As Integer
Adodc1.RecordSource = ""

Adodc1.CommandType = adCmdText

Adodc1.RecordSource = "SELECT * FROM `Cardapio_medidas`"

Adodc1.Refresh
Adodc1.Recordset.MoveLast
ultimoRegistrotitulo = Adodc1.Recordset("idCardapio_medidas").Value + 1
Form50.Text13 = ultimoRegistrotitulo

End Function

