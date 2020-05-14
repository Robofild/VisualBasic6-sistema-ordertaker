VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form402 
   Caption         =   "Form9"
   ClientHeight    =   4845
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10500
   Icon            =   "aTacressimos.frx":0000
   LinkTopic       =   "Form9"
   ScaleHeight     =   4845
   ScaleWidth      =   10500
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      DataField       =   "valor"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   8520
      TabIndex        =   4
      Top             =   960
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Usar"
      Enabled         =   0   'False
      Height          =   615
      Left            =   7320
      Picture         =   "aTacressimos.frx":0B02
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   7320
      Picture         =   "aTacressimos.frx":1504
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton Command15 
      Height          =   375
      Left            =   7800
      Picture         =   "aTacressimos.frx":1F06
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   480
      Width           =   375
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   480
      Top             =   5280
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
      RecordSource    =   "Acressimos"
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
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "aTacressimos.frx":2908
      Height          =   4260
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   7514
      _Version        =   393216
      Style           =   1
      ListField       =   "Acressimo"
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
      Caption         =   "Valor para acréssimo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8520
      TabIndex        =   5
      Top             =   720
      Width           =   2295
   End
End
Attribute VB_Name = "Form402"
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

Adodc1.RecordSource = "SELECT * FROM `Acressimos` WHERE `Acressimo` LIKE '" & nome & "'"

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
'DataCombo1.Text = ""
Call Command1_Click
End Sub

Private Sub DataCombo1_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
Call Command2_Click
End If

End Sub

Private Sub Form_Load()
'trarar erro
On Error GoTo error

Adodc1.RecordSource = ""

Adodc1.CommandType = adCmdText
                       
Adodc1.RecordSource = "SELECT * FROM `Acressimos` ORDER BY `Acressimos`.`Acressimo` ASC"
Adodc1.Refresh

Exit Sub

error:

Adodc1.RecordSource = ""

Adodc1.CommandType = adCmdText
                       
Adodc1.RecordSource = "SELECT * FROM `Acressimos` ORDER BY `Acressimos`.`Acressimo` ASC"
Adodc1.Refresh

Exit Sub

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

Adodc1.RecordSource = "SELECT * FROM `Acressimos` WHERE `Acressimo` LIKE '" & tipo & "'"

Adodc1.Refresh

 


    If (Adodc1.Recordset.EOF = True) Then
    
    ultimoRegistrotitulo
    Adodc1.Recordset.AddNew
  
    Adodc1.Recordset("Acressimo").Value = tipo
    Adodc1.Recordset.Update
      Form401.Text22.Text = Form401.Text3.Text & " " & Form401.Text5.Text & "  AC  " & tipo
    If Text1.Text <> "" Then
   Form401.Text16.Text = Text1.Text
    End If
    If Form401.Text7.Text <> "" Then
  
   Form401.Text7.Text = Form401.Text7.Text & "  AC  " & tipo
   Form401.Text22.Text = Form401.Text7.Text & " " & Form401.Text5.Text & "  AC  " & tipo
   Form401.Show
    Else
  MsgBox "Primeiro escolha o produto para receber o acréssimo", , "Produto não Escolhido"
    
    Form401.Show
    
    End If
    Else
              If Text1.Text <> "" Then
   Form401.Text16.Text = Text1.Text
    End If
      Form401.Text22.Text = Form401.Text3.Text + " " + Form401.Text5.Text & "  AC  " & tipo
              If Form401.Text7.Text <> "" Then
               If Text1.Text <> "" Then
   Form401.Text16.Text = Text1.Text
    End If
            Form401.Text7.Text = Form401.Text7.Text & "  AC  " & tipo
            Form401.Show
            End If
             End If
   ' Form50.Text5.SetFocus
    Unload Me

Else
'MsgBox "Crie ou selecione uma Categoria para seu cardápio", vbInformation, "Categoria não definida!!!"


End If


Form401.Timer2.Interval = 100

End Function

Public Function ultimoRegistrotitulo() As Integer
Adodc1.RecordSource = ""

Adodc1.CommandType = adCmdText

Adodc1.RecordSource = "SELECT * FROM `Acressimos`"

Adodc1.Refresh
Adodc1.Recordset.MoveLast
ultimoRegistrotitulo = Adodc1.Recordset("id").Value + 1
'Form50.Text11 = ultimoRegistrotitulo

End Function

Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = SoPonto(KeyAscii)

End Sub

Private Sub Text1_LostFocus()
 Text1.Text = Replace(Text1.Text, ",", ".")
'trarar erro
On Error GoTo error
If Text1.Text <> "" Then
Dim testo1 As Double
   
 
testo1 = CDbl(Text1.Text)
'Text1.Text = Format(Text1.Text, "Currency")
Adodc1.Recordset.Update
End If



Exit Sub

error:

Exit Sub
End Sub
