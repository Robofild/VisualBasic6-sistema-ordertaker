VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form51_4 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Crie ingredientes para "
   ClientHeight    =   3330
   ClientLeft      =   6600
   ClientTop       =   4530
   ClientWidth     =   8535
   Icon            =   "cadastroDeItens.frx":0000
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   8535
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      DataField       =   "nomeTipoIngrediente"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   360
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Remover"
      Height          =   615
      Left            =   4320
      Picture         =   "cadastroDeItens.frx":0A02
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   480
      Width           =   855
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   8160
      Top             =   3120
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
      RecordSource    =   "ingredientes_por_id_anexos"
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
   Begin MSDataListLib.DataList DataList1 
      Bindings        =   "cadastroDeItens.frx":1404
      Height          =   2700
      Left            =   5280
      TabIndex        =   4
      Top             =   360
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   4763
      _Version        =   393216
      ListField       =   "ingredientes_por_id_anexoscol"
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
   Begin VB.CommandButton Command2 
      Caption         =   "Usar"
      Enabled         =   0   'False
      Height          =   615
      Left            =   3000
      Picture         =   "cadastroDeItens.frx":1419
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   3000
      Picture         =   "cadastroDeItens.frx":1E1B
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   375
   End
   Begin VB.CommandButton Command15 
      Height          =   375
      Left            =   3480
      Picture         =   "cadastroDeItens.frx":281D
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   375
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "cadastroDeItens.frx":321F
      Height          =   2820
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Você pode e deve criar a medida que mais atende ao produto em questão!"
      Top             =   360
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   4974
      _Version        =   393216
      Style           =   1
      ListField       =   "nomeTipoIngrediente"
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
      Left            =   3120
      Top             =   2640
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
      RecordSource    =   "formIngredientes"
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
Attribute VB_Name = "Form51_4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim PossivelNovoIten As String
Dim IndicedoCardapi As Integer


Private Sub Command1_Click()
Dim nome As String

nome = Trim(DataCombo1.Text + "%")
Adodc1.RecordSource = ""

Adodc1.CommandType = adCmdText
                       
Adodc1.RecordSource = "SELECT * FROM `formIngredientes` WHERE `nomeTipoIngrediente` LIKE '" & nome & "' ORDER BY `nomeTipoIngrediente` ASC"
Adodc1.Refresh

 DataCombo1.Text = Format(PossivelNovoIten, ">")
If DataCombo1.Text <> "" Then
Command2.Enabled = True
End If
End Sub

Private Sub Command15_Click()
DataCombo1.Text = ""

Call Command1_Click
End Sub

Private Sub Command2_Click()
DataCombo1.Text = Format(DataCombo1.Text, ">")

naoIncluirDuplicidadeeee (DataCombo1.Text)
'Form51.Hide
End Sub

Private Sub Command3_Click()
IncluirdadoNalistaDeIgredientes
End Sub

Private Sub Command4_Click()
dELETitemlistaIngrediente

End Sub

Private Sub DataCombo1_Click(Area As Integer)

If DataCombo1.Text <> "" Then
Command2.Enabled = True
End If


End Sub

Private Sub DataCombo1_DblClick(Area As Integer)
transportaParaListadeItens
End Sub

Private Sub DataCombo1_LostFocus()
PossivelNovoIten = Format(DataCombo1.Text, ">")
'Text1.Text = Format(DataCombo1.Text, ">")
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
'If KeyAscii = vbKeyReturn Then
       
        SendKeys ("{TAB}")
        KeyAscii = 0
        End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
REVOLTARlistaAdodc2itensjaescolhidos
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
                       
Adodc1.RecordSource = "SELECT * FROM `formIngredientes` ORDER BY `formIngredientes`.`nomeTipoIngrediente` ASC"
Adodc1.Refresh

REVOLTARlistaAdodc2itensjaescolhidos


End Sub

Public Function naoIncluirDuplicidadeeee(produto As String)



If produto <> "" Then


ConServer

Dim sql As String
Dim rs As New ADODB.Recordset

Set rs = New ADODB.Recordset
Set rs.ActiveConnection = con

sql = "SELECT * FROM `formIngredientes` WHERE `nomeTipoIngrediente` LIKE '" & produto & "' ORDER BY `nomeTipoIngrediente` ASC "

rs.Open sql


    If (rs.EOF = False) Then
    
   
      transportaParaListadeItens
    Else
       IncluirdadoNalistaDeIgredientes
    End If
    
   
    ' entra com o item na lista do lado
    
    'Form50.Text13.Text = Adodc1.Recordset("idCardapio_medidas").Value
    'Form50.Text12.Text = DataCombo1.Text
    'Form50.Show
    
  
   Set rs = Nothing
    
    'Form50.Text4.SetFocus
    'Form51.Hide
    

Else
MsgBox "Crie ou selecione um item para seu produto", vbInformation, "Ingrediente  não definido!!!"


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



Public Sub transportaParaListadeItens()
 'captiar dados do cardapio
    'IndicedoCardapi = Form50.Text3.Text
    If Form50.Text3 <> "" Then
    IndicedoCardapi = Form50.Text3.Text
    Adodc2.Recordset.AddNew
    Adodc2.Recordset("chaveDaProdutp").Value = IndicedoCardapi
    Adodc2.Recordset("ingredientes_por_id_anexoscol").Value = DataCombo1.Text
    Adodc2.Recordset.Update
    Adodc2.Refresh
    Adodc2.Recordset.MoveLast
    End If
End Sub

Public Sub IncluirdadoNalistaDeIgredientes()
   'salvar o item na lista
    
    If DataCombo1.Text <> "" Then
   
    Adodc1.Recordset.AddNew
   Text1.Text = Format(DataCombo1.Text, ">")
    Adodc1.Recordset.Update
    Adodc1.Refresh
        'DataCombo1.Text = ""
    DataCombo1.SetFocus
    
    transportaParaListadeItens
    End If
End Sub

Public Sub dELETitemlistaIngrediente()
Dim PRODUTODEL As String
PRODUTODEL = DataList1.Text

ConServer

Dim sql As String
Dim rs As New ADODB.Recordset

Set rs = New ADODB.Recordset
Set rs.ActiveConnection = con

sql = "DELETE  FROM `ingredientes_por_id_anexos` WHERE `ingredientes_por_id_anexoscol` LIKE '" & PRODUTODEL & "' "

rs.Open sql
REVOLTARlistaAdodc2itensjaescolhidos


    
   Set rs = Nothing
End Sub

Public Sub REVOLTARlistaAdodc2itensjaescolhidos()

Adodc2.RecordSource = ""

Adodc2.CommandType = adCmdText
                       
Adodc2.RecordSource = "SELECT * FROM `ingredientes_por_id_anexos` WHERE `chaveDaProdutp` = '" & Form50.Text3.Text & "'"
Adodc2.Refresh

Form50.Adodc1.RecordSource = ""

Form50.Adodc1.CommandType = adCmdText
                       
Form50.Adodc1.RecordSource = "SELECT * FROM `ingredientes_por_id_anexos` WHERE `chaveDaProdutp` = '" & Form50.Text3.Text & "'"
Form50.Adodc1.Refresh

End Sub
