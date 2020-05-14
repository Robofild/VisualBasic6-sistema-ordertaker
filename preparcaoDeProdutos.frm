VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form51_7 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form9"
   ClientHeight    =   4110
   ClientLeft      =   7995
   ClientTop       =   5055
   ClientWidth     =   9075
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   9075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "preparcaoDeProdutos.frx":0000
      Height          =   2175
      Left            =   3840
      TabIndex        =   18
      Top             =   6960
      Visible         =   0   'False
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   3836
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
      Caption         =   "adodc2"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "preparcaoDeProdutos.frx":0015
      Height          =   1335
      Left            =   6120
      TabIndex        =   17
      Top             =   5280
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   2355
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
      Caption         =   "adodc3"
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
   Begin VB.CommandButton Command2 
      Caption         =   "&Salvar"
      Height          =   615
      Left            =   8040
      Picture         =   "preparcaoDeProdutos.frx":002A
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3360
      Width           =   855
   End
   Begin VB.TextBox Text3 
      DataField       =   "idingredientes_por_id_anexos"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   1200
      TabIndex        =   10
      Text            =   "Text3"
      Top             =   4440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text2 
      DataField       =   "chaveDaProdutp"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   4680
      TabIndex        =   9
      Text            =   "Text2"
      Top             =   5040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command15 
      Height          =   375
      Left            =   3120
      Picture         =   "preparcaoDeProdutos.frx":0A2C
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   840
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   3120
      Picture         =   "preparcaoDeProdutos.frx":142E
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   360
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   10680
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   3720
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox Text4 
      DataField       =   "ingrediente_instrucao"
      DataSource      =   "Adodc2"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3600
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "preparcaoDeProdutos.frx":1E30
      Top             =   360
      Width           =   5295
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Esta instrução ajuda o telemarketing  a concretizar a venda"
      Height          =   255
      Left            =   3600
      TabIndex        =   4
      Top             =   1800
      Width           =   5175
   End
   Begin VB.TextBox Text5 
      DataField       =   "obs"
      DataSource      =   "Adodc3"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3600
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "preparcaoDeProdutos.frx":1E36
      Top             =   2160
      Width           =   5295
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   615
      Left            =   2640
      TabIndex        =   2
      Top             =   4800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text6 
      DataField       =   "NumeroPrato"
      DataSource      =   "Adodc3"
      Height          =   285
      Left            =   10680
      TabIndex        =   1
      Text            =   "Text6"
      Top             =   3240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Height          =   615
      Left            =   9480
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "preparcaoDeProdutos.frx":1E3C
      Top             =   2400
      Visible         =   0   'False
      Width           =   2415
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   10560
      Top             =   1560
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      RecordSource    =   "ingrediente_instrucao_preparo"
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   375
      Left            =   6240
      Top             =   4560
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
      RecordSource    =   "ProntoAtendimentoTelemarket"
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
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "preparcaoDeProdutos.frx":1E42
      Height          =   2820
      Left            =   120
      TabIndex        =   12
      ToolTipText     =   "Você pode e deve criar a medida que mais atende ao produto em questão!"
      Top             =   360
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   4974
      _Version        =   393216
      Style           =   1
      BackColor       =   16777215
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
      Left            =   0
      Top             =   5400
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
   Begin VB.Label Label4 
      Caption         =   "51.7"
      Height          =   615
      Left            =   1320
      TabIndex        =   16
      Top             =   3600
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Label Label2 
      Caption         =   "Instrução de preparo"
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   0
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "t3"
      Height          =   255
      Left            =   1200
      TabIndex        =   14
      Top             =   4920
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Preparo Idependente por Produto"
      Height          =   255
      Left            =   3600
      TabIndex        =   13
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "Form51_7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()
selecioneTodos


End Sub
Public Sub pesquiseatributosParateemarket()
Dim NumContador As Integer


ConServer

Dim sql As String
Dim rs As New ADODB.Recordset

Set rs = New ADODB.Recordset
Set rs.ActiveConnection = con

sql = "SELECT * FROM `ProntoAtendimentoTelemarket` WHERE `NumeroPrato` =  '" & Text2.Text & "' "


rs.Open sql


    If (rs.EOF = False) Then
    'GIRAR O ADO PARA A POSICAO DO PRATO
    GIREPARAPOSICAOdopratotelemaaarketing
    
    Check1.Value = 1
    Text5.Enabled = True
'    Text5.SetFocus
    
       
        
    Else
     Text5.Text = ""
    Text5.Enabled = False
    'Adodc3.Recordset.AddNew
   
    
   End If

   Set rs = Nothing



End Sub
Public Sub selecioneTodos()
If Check1.Value = 1 Then
            DataCombo1.BackColor = &HFF0000
            Text5.Enabled = True
                If Text5.Text = "" Then
                  
                Adodc3.Recordset.AddNew
                Text6.Text = Val(Text2.Text)
                Text5.SetFocus
  




                End If

Else
'reatualize o prato

Text5.Text = ""
Text5.Enabled = False
DataCombo1.BackColor = &HFFFFFF
End If

End Sub


Private Sub Command1_Click()
Dim nome As String

nome = Trim(DataCombo1.Text + "%")
Adodc1.RecordSource = ""

Adodc1.CommandType = adCmdText
                       
Adodc1.RecordSource = "SELECT * FROM `ingredientes_por_id_anexos` WHERE `chaveDaProdutp` = '" & Text2.Text & "' AND `ingredientes_por_id_anexoscol` LIKE'" & nome & "'"
Adodc1.Refresh


End Sub

Private Sub Command15_Click()
DataCombo1.Text = ""
Call Command1_Click

End Sub

Private Sub Command2_Click()

 If Text4.Text <> "" Or Text5.Text <> "" Then
If Text4.Text <> "" Then
'chave do produto
 Adodc2.Recordset("cod_menu").Value = Text2.Text
 Adodc2.Recordset("cod_index").Value = Text3.Text
 'trarar erro
On Error GoTo error

Adodc2.Recordset.Update
Adodc2.Recordset.MoveFirst
Adodc2.Refresh
End If
    If Text5.Text <> "" Then
  'trarar erro
On Error GoTo errorRecord3

        Adodc3.Recordset("NumeroPrato").Value = Text2.Text
        Adodc3.Recordset("obs").Value = Text5.Text
          Adodc3.Recordset.Update
'       Adodc3.Recordset.MoveNext
        Adodc3.Refresh
    End If
    MsgBox "Salvo preparo com sucesso", vbYes, "Salvo com Sucesso"
Exit Sub

errorRecord3:
'             Adodc3.Refresh
        MsgBox " Salvo preparo com sucesso !", vbYes, "Salvo com Sucesso"
                Exit Sub




error:
             Adodc2.Refresh
            
             'Adodc3.Recordset("NumeroPrato").Value = Text2.Text
             If Text5.Text <> "" Then
              Adodc3.Recordset.Update
              Adodc3.Recordset.MoveFirst
              Adodc3.Refresh
            End If
        MsgBox "ib  Salvo preparo com sucesso", vbYes, "Salvo com Sucesso"
                Exit Sub


MsgBox "i 3 Salvo preparo com sucesso", vbYes, "Salvo com Sucesso"

End If

MsgBox "i 4Salvo preparo com sucesso", vbYes, "Salvo com Sucesso"
Exit Sub
MsgBox "i 5Salvo preparo com sucesso", vbYes, "Salvo com Sucesso"
End Sub

Private Sub DataCombo1_Click(Area As Integer)
 retornaoiddoproduto (DataCombo1.Text)
 verificarSeharegistroparraesteProcesso



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
     
     Adodc2.RecordSource = ""

    Adodc2.CommandType = adCmdText

Adodc2.RecordSource = "SELECT * FROM `ingredientes_por_id_anexos` WHERE `idingredientes_por_id_anexos` = '" & Text3.Text & "' AND `chaveDaProdutp` = '" & Text2.Text & "'"

Adodc2.Refresh

     'todo
     'Adodc3.RecordSource = ""

    'Adodc3.CommandType = adCmdText

'Adodc2.RecordSource = "SELECT * FROM `ingredientes_por_id_anexos` WHERE `idingredientes_por_id_anexos` = '" & Text3.Text & "' AND `chaveDaProdutp` = '" & Text2.Text & "'"

'Adodc2.Refresh
     
     
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

'Private Sub DataCombo2_Change()
'If DataCombo2.Text <> "" Then
'Command2.Enabled = True
'End If
'End Sub

'Private Sub DataCombo2_LostFocus()
'naoIncluirDuplicidadeeee (DataCombo2.Text)
'End Sub



 

Public Function verificarSeharegistroparraesteProcesso()
Dim CHavedoProduto As Integer

 Dim IngredientesPorIdeAnexos As Integer





ConServer

Dim sql As String
Dim rs As New ADODB.Recordset

Set rs = New ADODB.Recordset
Set rs.ActiveConnection = con

sql = "SELECT * FROM `ingrediente_instrucao_preparo` WHERE `cod_menu` = '" & Text2.Text & "' AND `cod_index` = '" & Text3.Text & "'"

rs.Open sql


    If (rs.EOF = False) Then
    movertestomemorandoparaOregistro
    
    
    
    
   ' Call Command2_Click
   
      'nada acontece
    Else
    'trarar erro
On Error GoTo error

CHavedoProduto = Adodc2.Recordset("chaveDaProdutp").Value

 IngredientesPorIdeAnexos = Adodc2.Recordset("idingredientes_por_id_anexos").Value
Adodc2.RecordSource = ""
 Adodc2.CommandType = adCmdText

Adodc2.RecordSource = "SELECT * FROM `ingrediente_instrucao_preparo` ORDER BY `cod_index` ASC"

Adodc2.Refresh
 Adodc2.Recordset.AddNew
       
       
 Adodc2.Recordset("cod_menu").Value = CHavedoProduto
Adodc2.Recordset("cod_index").Value = IngredientesPorIdeAnexos

     '  Text4.Text = DataCombo2.Text
      ' Adodc2.Recordset.Update
       'Adodc2.Recordset.MoveFirst
      ' Adodc2.Refresh
       
       'incluir na lista
    End If
    
Exit Function
    ' entra com o item na lista do lado
    
    'Form50.Text13.Text = Adodc1.Recordset("idCardapio_medidas").Value
    'Form50.Text12.Text = DataCombo1.Text
    'Form50.Show
    
  
    
    'Form50.Text4.SetFocus
    'Form51.Hide
    

'1sgBox "Crie ou selecione um medida para receita do produto", vbInformation, "Medida do ingrediente não definido!!!"



error:

   Set rs = Nothing
Exit Function




End Function


Public Sub movertestomemorandoparaOregistro()
    Adodc2.RecordSource = ""

    Adodc2.CommandType = adCmdText

Adodc2.RecordSource = "SELECT * FROM `ingrediente_instrucao_preparo` WHERE `cod_menu` = '" & Text2.Text & "' AND `cod_index` = '" & Text3.Text & "'"
Adodc2.Refresh
If Adodc2.Recordset.BOF = False Then

End If

End Sub



Public Sub GIREPARAPOSICAOdopratotelemaaarketing()

Adodc3.RecordSource = ""

Adodc3.CommandType = adCmdText
                       
Adodc3.RecordSource = "SELECT * FROM `ProntoAtendimentoTelemarket` WHERE `NumeroPrato` =  '" & Text2.Text & "' "
Adodc3.Refresh

 
End Sub

Private Sub DataCombo1_DblClick(Area As Integer)
Text4.Enabled = True
Text4.SetFocus
End Sub

Private Sub Form_Load()
Text4.Text = ""
End Sub

Private Sub Text2_Change()

    pesquiseatributosParateemarket
End Sub

Private Sub Text4_LostFocus()
Text4.Text = Format(Text4.Text, ">")

End Sub

Private Sub Text5_LostFocus()
If Text5 = "" Then
Check1.Value = 0
Text5.Enabled = False
Adodc3.Refresh
Else
Text5.Text = Format(Text5.Text, ">")

'SalvarAjudaTelemarketing
End If
End Sub


Public Sub SalvarAjudaTelemarketing()


ConServer

Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset
Set rs.ActiveConnection = con
sql = "SELECT * FROM `ProntoAtendimentoTelemarket` WHERE `NumeroPrato` = '" & Text2.Text & "'"
rs.Open sql

If rs.BOF = True Then
rs.Close
sql = "INSERT INTO `ProntoAtendimentoTelemarket` ( `NumeroPrato`, `obs`) VALUES ( '" & Text2.Text & "', '" & Text5.Text & "' )"
'rs.Open sql
Else
rs.Close

sql = "UPDATE `ProntoAtendimentoTelemarket` SET `obs` =  '" & Text5.Text & "'WHERE `ProntoAtendimentoTelemarket`.`NumeroPrato` =  '" & Text2.Text & "'"

End If
rs.Open sql

'rs.Close sql



 Set rs = Nothing



End Sub
