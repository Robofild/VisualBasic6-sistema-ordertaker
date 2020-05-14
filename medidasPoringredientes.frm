VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form51_5 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Crie ingredientes para "
   ClientHeight    =   5940
   ClientLeft      =   7050
   ClientTop       =   4530
   ClientWidth     =   12960
   Icon            =   "medidasPoringredientes.frx":0000
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   12960
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text5 
      BackColor       =   &H00C0FFC0&
      DataField       =   "Quantidade_num_id_anexo"
      DataSource      =   "Adodc3"
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
      Left            =   3240
      TabIndex        =   1
      Text            =   "Text5"
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Salvar"
      Enabled         =   0   'False
      Height          =   615
      Left            =   6720
      Picture         =   "medidasPoringredientes.frx":0A02
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1440
      Width           =   855
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   375
      Left            =   9360
      Top             =   1200
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
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
   Begin VB.TextBox Text4 
      DataField       =   "medida_por_igrediente"
      DataSource      =   "Adodc2"
      Height          =   375
      Left            =   7560
      TabIndex        =   10
      Text            =   "Text4"
      Top             =   4080
      Visible         =   0   'False
      Width           =   3015
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Bindings        =   "medidasPoringredientes.frx":1404
      DataField       =   "listaInsertIngrediente"
      DataSource      =   "Adodc3"
      Height          =   2820
      Left            =   4320
      TabIndex        =   2
      Top             =   600
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   4974
      _Version        =   393216
      Style           =   1
      ListField       =   "medida_por_igrediente"
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
      DataField       =   "idingredientes_por_id_anexos"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   11160
      TabIndex        =   8
      Text            =   "Text3"
      Top             =   2280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text2 
      DataField       =   "chaveDaProdutp"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   9600
      TabIndex        =   7
      Text            =   "Text2"
      Top             =   1920
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command15 
      Height          =   375
      Left            =   7200
      Picture         =   "medidasPoringredientes.frx":1419
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   600
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   6720
      Picture         =   "medidasPoringredientes.frx":1E1B
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   2640
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   4800
      Top             =   4320
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
      RecordSource    =   "MedidasporIngredientes"
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
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "medidasPoringredientes.frx":281D
      Height          =   2820
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   "Você pode e deve criar a medida que mais atende ao produto em questão!"
      Top             =   600
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   4974
      _Version        =   393216
      Style           =   1
      ListField       =   "ingredientes_por_id_anexoscol"
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
      Left            =   480
      Top             =   4200
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
      RecordSource    =   "ingredientes_por_id_anexos"
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
   Begin VB.Label Label3 
      Caption         =   "Quantidade"
      Height          =   255
      Left            =   3240
      TabIndex        =   12
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Escolha o produto para criar uma medida"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   240
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "Entre com a medida da receita"
      Height          =   255
      Left            =   4320
      TabIndex        =   9
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "Form51_5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim nome As String

nome = Trim(DataCombo2.Text + "%")
Adodc2.RecordSource = ""

Adodc2.CommandType = adCmdText
                       
Adodc2.RecordSource = "SELECT * FROM `MedidasporIngredientes` WHERE `medida_por_igrediente` LIKE'" & nome & "'"
Adodc2.Refresh

 
If DataCombo2.Text <> "" Then
Command2.Enabled = True
End If
End Sub

Private Sub Command15_Click()
DataCombo2.Text = ""
Call Command1_Click

End Sub

Private Sub Command2_Click()
If DataCombo2.Text <> "" And Text5 <> "" Then


Adodc3.Recordset.Update

Adodc3.Recordset.MoveFirst
Text5.Text = ""
Adodc3.Refresh
MsgBox " O produto '" & DataCombo1.Text & "' foi salvo com a medida '" & DataCombo2.Text & "' ", vbYes, "Salvo com sucesso!"
DataCombo1.SetFocus

Else
MsgBox "Complete os dados antes de salvar ", , "Não salvo"
End If


Exit Sub
End Sub

Private Sub DataCombo1_Click(Area As Integer)
 retornaoiddoproduto (DataCombo1.Text)

  

End Sub

Public Sub retornaoiddoproduto(textselecionado)

ConServer

Dim sql As String
Dim rs As New ADODB.Recordset

Set rs = New ADODB.Recordset
Set rs.ActiveConnection = con

sql = "SELECT `idingredientes_por_id_anexos` FROM `ingredientes_por_id_anexos` WHERE `ingredientes_por_id_anexoscol` LIKE '" & textselecionado & "' AND `chaveDaProdutp` = '" & Text2.Text & "'"

rs.Open sql


    If (rs.EOF = False) Then
    
   
     Text3.Text = rs.Fields("idingredientes_por_id_anexos").Value
     
     Adodc3.RecordSource = ""

    Adodc3.CommandType = adCmdText

Adodc3.RecordSource = "SELECT * FROM `ingredientes_por_id_anexos` WHERE `idingredientes_por_id_anexos` = '" & Text3.Text & "' AND `chaveDaProdutp` = '" & Text2.Text & "'"

Adodc3.Refresh

     
     
     
    Else
       
    End If
    
   
    ' entra com o item na lista do lado
    
    '
    'Form50.Text12.Text = DataCombo1.Text
    'Form50.Show
    
  
   Set rs = Nothing
End Sub

Public Sub movaadodc3paraposiçaosalutar()
   
   
    '
   
End Sub

Private Sub DataCombo2_Change()
If DataCombo2.Text <> "" Then
Command2.Enabled = True
End If
End Sub

Private Sub DataCombo2_LostFocus()
naoIncluirDuplicidadeeee (DataCombo2.Text)
End Sub
Public Function naoIncluirDuplicidadeeee(produto As String)



If produto <> "" Then


ConServer

Dim sql As String
Dim rs As New ADODB.Recordset

Set rs = New ADODB.Recordset
Set rs.ActiveConnection = con

sql = "SELECT * FROM `MedidasporIngredientes` WHERE `medida_por_igrediente` LIKE '" & produto & "'"

rs.Open sql


    If (rs.EOF = False) Then
    Call Command2_Click
   
      'nada acontece
    Else
       Adodc2.Recordset.AddNew
       Text4.Text = DataCombo2.Text
       Adodc2.Recordset.Update
       Adodc2.Recordset.MoveFirst
       Adodc2.Refresh
       
       'incluir na lista
    End If
    
   
    ' entra com o item na lista do lado
    
    'Form50.Text13.Text = Adodc1.Recordset("idCardapio_medidas").Value
    'Form50.Text12.Text = DataCombo1.Text
    'Form50.Show
    
  
   Set rs = Nothing
    
    'Form50.Text4.SetFocus
    'Form51.Hide
    

Else
MsgBox "Crie ou selecione um medida para receita do produto", vbInformation, "Medida do ingrediente não definido!!!"


End If




End Function

Private Sub Text5_Change()
DataCombo2.Text = Text5.Text


End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
KeyAscii = (SoNumeros(KeyAscii))
    If KeyAscii = 0 Then
    End If
End Sub

Private Sub Text5_LostFocus()
Call Command1_Click
End Sub
