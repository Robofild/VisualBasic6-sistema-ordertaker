VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form4440 
   Caption         =   "Personalize botão de atendimento"
   ClientHeight    =   5445
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7860
   Icon            =   "personalizeNewRrec.frx":0000
   LinkTopic       =   "Form9"
   ScaleHeight     =   5445
   ScaleWidth      =   7860
   Begin VB.CommandButton Command1 
      Caption         =   "Personalize"
      Height          =   2055
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      Caption         =   "Personalize"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   3960
      TabIndex        =   18
      Top             =   360
      Width           =   3495
      Begin VB.CommandButton Command3 
         Caption         =   "Fonte"
         Height          =   495
         Left            =   360
         TabIndex        =   23
         Top             =   2040
         Width           =   2895
      End
      Begin VB.CommandButton cmdSalvar 
         Caption         =   "Salvar"
         Height          =   615
         Left            =   1440
         Picture         =   "personalizeNewRrec.frx":0A02
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   3720
         Width           =   855
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Imagem"
         Height          =   615
         Left            =   360
         TabIndex        =   21
         Top             =   2760
         Width           =   2895
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cores"
         Height          =   495
         Left            =   360
         TabIndex        =   19
         Top             =   1320
         Width           =   2895
      End
      Begin MSAdodcLib.Adodc Adodc3 
         Height          =   375
         Left            =   480
         Top             =   120
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   4471
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
         RecordSource    =   "Cardapio"
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
         Bindings        =   "personalizeNewRrec.frx":1404
         Height          =   420
         Left            =   360
         TabIndex        =   20
         Top             =   600
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   741
         _Version        =   393216
         ListField       =   "tipo"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "Texto"
         Height          =   375
         Left            =   1560
         TabIndex        =   24
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.TextBox Text2 
      Height          =   405
      Left            =   8160
      TabIndex        =   17
      Text            =   "Text2"
      Top             =   480
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      DataField       =   "Color"
      DataSource      =   "Adodc2"
      Height          =   375
      Left            =   8160
      TabIndex        =   16
      Text            =   "Text3"
      Top             =   1560
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text4 
      DataField       =   "texto"
      DataSource      =   "Adodc2"
      Height          =   375
      Left            =   8160
      TabIndex        =   15
      Text            =   "Text4"
      Top             =   2160
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text5 
      DataField       =   "fonte0"
      DataSource      =   "Adodc2"
      Height          =   375
      Left            =   8160
      TabIndex        =   14
      Text            =   "Text5"
      Top             =   2760
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text6 
      DataField       =   "fonte1"
      DataSource      =   "Adodc2"
      Height          =   375
      Left            =   8160
      TabIndex        =   13
      Text            =   "Text6"
      Top             =   3360
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text7 
      DataField       =   "fonte2"
      DataSource      =   "Adodc2"
      Height          =   375
      Left            =   8160
      TabIndex        =   12
      Text            =   "Text7"
      Top             =   3840
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text8 
      DataField       =   "fonte3"
      DataSource      =   "Adodc2"
      Height          =   375
      Left            =   8160
      TabIndex        =   11
      Text            =   "Text8"
      Top             =   4440
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Text9 
      DataField       =   "fonte4"
      DataSource      =   "Adodc2"
      Height          =   375
      Left            =   8160
      TabIndex        =   10
      Text            =   "Text9"
      Top             =   5040
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Text10 
      DataField       =   "fonte5"
      DataSource      =   "Adodc2"
      Height          =   285
      Left            =   8160
      TabIndex        =   9
      Text            =   "Text10"
      Top             =   5640
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      DataField       =   "imagemCaminho"
      DataSource      =   "Adodc2"
      Height          =   285
      Left            =   10560
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   2160
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text11 
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Text            =   "Text11"
      Top             =   6120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   375
      Left            =   3960
      TabIndex        =   5
      Top             =   7920
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Text12 
      DataField       =   "IndiceProtocoloCardapio"
      DataSource      =   "Adodc2"
      Height          =   375
      Left            =   10560
      TabIndex        =   4
      Text            =   "Text12"
      Top             =   3000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text13 
      Height          =   495
      Left            =   10560
      TabIndex        =   2
      Text            =   "Text13"
      Top             =   480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text14 
      DataField       =   "nomeOriginaldoBtn"
      DataSource      =   "Adodc2"
      Height          =   375
      Left            =   8160
      TabIndex        =   1
      Text            =   "Text14"
      Top             =   1080
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Text15 
      DataField       =   "CardapioOficial"
      DataSource      =   "Adodc2"
      Height          =   375
      Left            =   10560
      TabIndex        =   0
      Text            =   "Text15"
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "personalizeNewRrec.frx":1419
      Height          =   1455
      Left            =   6120
      TabIndex        =   3
      Top             =   7440
      Visible         =   0   'False
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   2566
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "personalizeNewRrec.frx":142E
      Height          =   5175
      Left            =   12240
      TabIndex        =   6
      Top             =   840
      Visible         =   0   'False
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   9128
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
      Caption         =   "ligado ao adodc3 dever apresentar os tipos de acordo com o titulo selecionado"
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   7920
      Top             =   6120
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3000
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000006&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Left            =   720
      Top             =   1440
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Protocolo cardapio"
      Height          =   375
      Left            =   10560
      TabIndex        =   26
      Top             =   2760
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "Form4440"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim imagem As String
Dim EndImagem As String
Dim oficialName As String
Dim VNomebtn As String
Dim vColor As String
Dim vText As String
Dim vFonte(5) As String
Dim ProtocoloOficial As Integer
Dim ChavedeSalvamento As Integer





Private Sub cmdSalvar_Click()
'Adodc1.Recordset.AddNew
passartodos (ChavedeSalvamento)
'trarar erro
On Error GoTo error

Adodc2.Recordset.Update
Adodc2.Recordset.MoveFirst
Adodc2.Refresh

MsgBox "Personalizado com sucesso o botão", , "Sucesso!"
Exit Sub

error:
'Adodc2.Refresh
Exit Sub
MsgBox "Personalizado com sucesso o botão", , "Sucesso!"

End Sub

Private Sub Command2_Click()
Command1.Enabled = False
Command1.BackColor = colorOPT

Text3.Text = Command1.BackColor
vColor = Command1.BackColor
Command1.Enabled = True
End Sub

Private Sub Command3_Click()


Form444.CommonDialog1.ShowFont

vFonte(0) = Form444.CommonDialog1.FontName
vFonte(1) = Form444.CommonDialog1.FontSize
vFonte(2) = Form444.CommonDialog1.FontItalic
vFonte(3) = Form444.CommonDialog1.FontBold
vFonte(4) = Form444.CommonDialog1.FontStrikethru
vFonte(5) = Form444.CommonDialog1.FontUnderline
'case fonte igual a null
Text5.Text = Form444.CommonDialog1.FontName
Text6.Text = Form444.CommonDialog1.FontSize
Text7.Text = Form444.CommonDialog1.FontItalic
Text8.Text = Form444.CommonDialog1.FontBold
Text9.Text = Form444.CommonDialog1.FontStrikethru
Text10.Text = Form444.CommonDialog1.FontUnderline


'trarar erro
On Error GoTo error

Command1.FontName = vFonte(0)
Command1.FontSize = vFonte(1)
Command1.FontItalic = vFonte(2)
Command1.FontBold = vFonte(3)
Command1.FontStrikethru = vFonte(4)
Command1.FontUnderline = vFonte(5)
Exit Sub
error:
Command3.SetFocus

End Sub

Private Sub Command4_Click()
If Command4.Caption = "Imagem" Then
                Dim imagem As String
            
                If Text1.Text <> "" Then
                Command4.Caption = "Remover imagem"
                End If
                CommonDialog1.InitDir = App.Path
                        CommonDialog1.FileName = ""
                        'CommonDialog1.Filter = "JPEG Image (*.jpg)|*.jpg|All Files (*.*)|*.*"
                         CommonDialog1.Filter = "JPEG Image (*.jpg)|*.jpg"
                        CommonDialog1.DialogTitle = "Open Image"
                        CommonDialog1.ShowOpen
                 
                        If CommonDialog1.FileName <> "" Then
                            Command1.Picture = LoadPicture(CommonDialog1.FileName)
                            imagem = LoadPicture(CommonDialog1.FileName)
                            EndImagem = (CommonDialog1.FileName)
                            Text1.Text = (CommonDialog1.FileName)
                            Else
                            Command1.Picture = LoadPicture()
                            imagem = LoadPicture()
                            EndImagem = ("")
                            End If
Else

           
                            imagem = ""
                            EndImagem = ""
                            Text1.Text = ""
                            Command4.Caption = "Imagem"
                            Command1.Picture = LoadPicture()
                            'retificarImagem

End If

End Sub



Private Sub DataCombo1_Click(Area As Integer)
Command1.Caption = DataCombo1.Text


End Sub

Private Sub Form_Load()
'oficialName =
'ProtocoloOficial = Text12.Text
'Alterarounovo




End Sub

Private Sub DataCombo1_LostFocus()
Text4.Text = DataCombo1.Text
vText = DataCombo1.Text
Command1.Caption = DataCombo1.Text
VNomebtn = DataCombo1.Text
End Sub

Public Sub passartodos(ChavedeSalvamento)


'trarar erro
'On Error GoTo error
If DataCombo1.Text <> "" Then
If ChavedeSalvamento = 1 Then
Text2.Text = oficialName
Text3.Text = vColor
Text4.Text = vText
Text1.Text = EndImagem
Text12.Text = ProtocoloOficial

Text5.Text = vFonte(0)
Text6.Text = vFonte(1)
Text7.Text = vFonte(2)
Text8.Text = vFonte(3)
Text9.Text = vFonte(4)
Text10.Text = vFonte(5)
Else


End If
Else
MsgBox "Não será possivel salvar estas alterações você tem que escolher o nome para o botão", vbYes, "Nome para o Botão?"
DataCombo1.SetFocus
End If
'Exit Sub
'error:
'Command2.BackColor = vbBlue
'Exit Sub



End Sub

Public Sub RetorNaosTiposPossiveis(titulo As Integer)
Dim resp As Integer
Form4440.Adodc3.RecordSource = ""

Form4440.Adodc3.CommandType = adCmdText

Form4440.Adodc3.RecordSource = "SELECT DISTINCT `tipo` FROM `Cardapio` WHERE `codTitulo` = '" & titulo & "'"

Form4440.Adodc3.Refresh
If Adodc3.Recordset.BOF = True Then
MsgBox "Você ainda não criou este cardápio ,não será possivel criar botões para ele ", , "Cardápio não Criado"
Unload Me
resp = MsgBox("Deseja criar agora esse cardapio?", vbYesNo, "Criar agora?")
        If resp = 6 Then
        Form444.Hide
        Form50.Show
        
        End If
Else
desbloquearControles
End If


End Sub


Public Sub Alterarounovo(indicePrimarykey As Integer)
Adodc2.RecordSource = ""

Adodc2.CommandType = adCmdText
    
        Adodc2.RecordSource = "SELECT * FROM `configuracao_tela_pedido` WHERE `idconfiguracao_tela_pedido` =  '" & indicePrimarykey & "'"

Adodc2.Refresh

If Adodc2.Recordset.BOF = False Then
ChavedeSalvamento = 0
transmitirParaAlteracao

SetAsconfigEncontradasnoButton
Else
Adodc2.Recordset.AddNew
ChavedeSalvamento = 1
End If

End Sub



Public Sub SetAsconfigEncontradasnoButton()
If Adodc2.Recordset("texto").Value <> "" Then
desbloquearControles
Command1.Caption = Adodc2.Recordset("texto").Value
DataCombo1.Text = Adodc2.Recordset("texto").Value
Else
bloquearControles
End If
'fonte
If Adodc2.Recordset("fonte0").Value <> " " And Adodc2.Recordset("fonte0").Value <> "" Then
Command1.FontName = Adodc2.Recordset("fonte0").Value
Command1.FontSize = Adodc2.Recordset("fonte1").Value
Command1.FontItalic = Adodc2.Recordset("fonte2").Value
Command1.FontBold = Adodc2.Recordset("fonte3").Value
Command1.FontStrikethru = Adodc2.Recordset("fonte4").Value
Command1.FontUnderline = Adodc2.Recordset("fonte5").Value
End If
'cor
If Adodc2.Recordset("Color").Value <> "" Then
Command1.BackColor = Adodc2.Recordset("Color").Value
End If
'imagem
If Adodc2.Recordset("imagemCaminho").Value <> "" Then
Command1.Picture = LoadPicture(Adodc2.Recordset("imagemCaminho").Value)
End If
Command1.Caption = Text4.Text
DataCombo1.Text = Text4.Text
'DataCombo1.SetFocus

End Sub

Public Sub transmitirParaAlteracao()



'trarar erro
On Error GoTo error

 oficialName = Text2.Text
 vColor = Text3.Text
 vText = Text4.Text
 Command1.Caption = Text4.Text
 EndImagem = Text1.Text
 ProtocoloOficial = Text13.Text

 vFonte(0) = Text5.Text
 vFonte(1) = Text6.Text
 vFonte(2) = Text7.Text
 vFonte(3) = Text8.Text
 vFonte(4) = Text9.Text
 vFonte(5) = Text9.Text


Exit Sub

error:

Exit Sub


End Sub

Public Function verificaSeAcorrespondenteAesteBotao(ByVal Nomebtn As String, ByVal numCardapio As Integer) As Boolean
ConServer
Dim IndicePrimario As Integer
Dim sql As String
Dim rs As New ADODB.Recordset

Set rs = New ADODB.Recordset
Set rs.ActiveConnection = con

sql = "SELECT * FROM `configuracao_tela_pedido` WHERE `nomeOriginaldoBtn` LIKE '" & Nomebtn & "'  AND `CardapioOficial` = '" & numCardapio & "'"

rs.Open sql

If rs.BOF = False Then
verificaSeAcorrespondenteAesteBotao = True
IndicePrimario = rs.Fields("idconfiguracao_tela_pedido").Value

Alterarounovo (IndicePrimario)


Else
Adodc2.Recordset.AddNew
verificaSeAcorrespondenteAesteBotao = False

End If

'rs.Close sql
 Set rs = Nothing
 Set rs = Nothing








End Function
Public Sub Personalize(numIdbuscar As Integer)


Adodc2.RecordSource = ""

Adodc2.CommandType = adCmdText

Adodc2.RecordSource = "SELECT * FROM `configuracao_tela_pedido` WHERE `idconfiguracao_tela_pedido` = ' " & numIdbuscar & " '"

Adodc2.Refresh

Command2.Caption = Adodc2.Recordset("texto").Value
'fonte
If Adodc2.Recordset("fonte0").Value <> "" Then
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
If Adodc2.Recordset("Color").Value <> "" Then
Command2.Picture = LoadPicture(Adodc2.Recordset("imagemCaminho").Value)
End If







End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)


'trarar erro
On Error GoTo error
Adodc2.Recordset.Update
Form4440.DataCombo1.SetFocus
Exit Sub

error:

Exit Sub

Form4440.DataCombo1.SetFocus
End Sub

Private Sub Text1_Change()
If Text1.Text <> "" Then
          Command4.Caption = "Remover imagem"
End If
End Sub

Private Sub Text13_Change()
Dim resp As Integer

Form4440.Adodc3.RecordSource = ""

Form4440.Adodc3.CommandType = adCmdText

Form4440.Adodc3.RecordSource = "SELECT DISTINCT `tipo` FROM `Cardapio` WHERE `codTitulo` = '" & Text13.Text & "'"

Form4440.Adodc3.Refresh

If Adodc3.Recordset.BOF = True Then
MsgBox "Você ainda não criou este cardápio ,não será possivel criar botões para ele ", , "Cardápio não Criado"
Unload Me
resp = MsgBox("Deseja criar agora esse cardapio?", vbYesNo, "Criar agora?")
        If resp = 6 Then
        Form444.Hide
        Form50.Show
        
        End If
Else
desbloquearControles
End If

End Sub

Private Sub Text2_Change()
'Text14.Text = Text2.Text
Dim X As String, Y As Integer

  X = Text2.Text
  Y = Val(Text13.Text)
  
If verificaSeAcorrespondenteAesteBotao(X, Y) Then
'trazer a tona configuaçoes setadas ateriormente
      
'fazer somente o complemento do que foi mudado





Else


'revespara nome do botao
oficialName = Text2.Text
Text14.Text = Text2.Text
Text15.Text = Text13.Text

'crie um novo parra o botão


End If
Exit Sub

End Sub


Public Sub SetReversoConfiguracao()


End Sub

Private Sub Text4_Change()
Command1.Caption = Text4.Text
End Sub


Public Sub desbloquearControles()
DataCombo1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
cmdSalvar.Enabled = True

End Sub

Public Sub bloquearControles()

DataCombo1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
cmdSalvar.Enabled = False

End Sub
