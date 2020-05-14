VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form403 
   Caption         =   "Form9"
   ClientHeight    =   6855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12030
   Icon            =   "atObservacoes.frx":0000
   LinkTopic       =   "Form9"
   ScaleHeight     =   6855
   ScaleWidth      =   12030
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command15 
      Height          =   375
      Left            =   7440
      Picture         =   "atObservacoes.frx":0A02
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   6960
      Picture         =   "atObservacoes.frx":1404
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Usar"
      Enabled         =   0   'False
      Height          =   615
      Left            =   6960
      Picture         =   "atObservacoes.frx":1E06
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   960
      Width           =   855
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "atObservacoes.frx":2808
      Height          =   4260
      Left            =   240
      TabIndex        =   3
      Top             =   0
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   7514
      _Version        =   393216
      Style           =   1
      ListField       =   "observacao"
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
      Left            =   1800
      Top             =   4800
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
      RecordSource    =   "at_oBservacao"
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
End
Attribute VB_Name = "Form403"
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

Adodc1.RecordSource = "SELECT * FROM `at_oBservacao` WHERE `observacao` LIKE '" & nome & "'"

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
Form402.Hide
End Sub

Private Sub DataCombo1_Click(Area As Integer)

If DataCombo1.Text <> "" Then
Command2.Enabled = True
End If


End Sub

Private Sub DataCombo1_DblClick(Area As Integer)
DataCombo1.Text = ""
End Sub

Private Sub Form_Load()
Adodc1.RecordSource = ""

Adodc1.CommandType = adCmdText
                       
Adodc1.RecordSource = "SELECT * FROM `at_oBservacao` ORDER BY `at_oBservacao`.`observacao` ASC"
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
'Form50.Text5.SetFocus
End Sub
Public Function naoIncluirDuplicidadeeee(tipo As String)



If tipo <> "" Then
tipo = Format(tipo, ">")

Adodc1.RecordSource = ""

Adodc1.CommandType = adCmdText

Adodc1.RecordSource = "SELECT * FROM `at_oBservacao` WHERE `observacao` LIKE '" & tipo & "'"

Adodc1.Refresh

 


    If (Adodc1.Recordset.EOF = True) Then
    
    ultimoRegistrotitulo
    Adodc1.Recordset.AddNew
  
    Adodc1.Recordset("observacao").Value = tipo
    Adodc1.Recordset.Update
      Form404.Text2.Text = Form401.Text3.Text & "  OBS  " & tipo & Form404.Text2.Text
    If Form401.Text7.Text <> "" And Form401.Text7.Visible = True Then
    
   Form401.Text13.Text = "  ##  " & tipo
   Form401.Show
    Else
  MsgBox "Primeiro escolha o produto para receber o acréssimo", , "Produto não Escolhido"
    
    Form401.Show
    
    End If
    Else
              If Form401.Text7.Text <> "" Then
             
            Form401.Text13.Text = "  # " & tipo
            Form401.Show
            End If
             End If
   ' Form50.Text5.SetFocus
    Unload Me

Else
'MsgBox "Crie ou selecione uma Categoria para seu cardápio", vbInformation, "Categoria não definida!!!"


End If




End Function

Public Function ultimoRegistrotitulo() As Integer
Adodc1.RecordSource = ""

Adodc1.CommandType = adCmdText

Adodc1.RecordSource = "SELECT * FROM `at_oBservacao`"

Adodc1.Refresh
Adodc1.Recordset.MoveLast
ultimoRegistrotitulo = Adodc1.Recordset("id").Value + 1
'Form50.Text11 = ultimoRegistrotitulo

End Function


