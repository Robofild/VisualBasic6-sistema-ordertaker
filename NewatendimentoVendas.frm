VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.ocx"
Begin VB.Form Form440 
   Caption         =   "Atendimento em andamento"
   ClientHeight    =   10080
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   19080
   Icon            =   "NewatendimentoVendas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form9"
   ScaleHeight     =   10080
   ScaleWidth      =   19080
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text9 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12600
      TabIndex        =   21
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4560
      TabIndex        =   20
      Top             =   1560
      Visible         =   0   'False
      Width           =   9495
   End
   Begin VB.Frame Frame1 
      Caption         =   "Compra em andamento"
      Height          =   7935
      Left            =   20760
      TabIndex        =   16
      Top             =   1320
      Width           =   6855
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   840
         TabIndex        =   17
         Text            =   "Text4"
         Top             =   4440
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Sub Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   18
         Top             =   4200
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command7 
      Caption         =   "F8                                                                      Ver Pedido                                             "
      Height          =   975
      Left            =   16920
      TabIndex        =   15
      Top             =   7920
      Width           =   2535
   End
   Begin VB.CommandButton Command6 
      Caption         =   "F6                                                                       Meia ½                                              "
      Height          =   975
      Left            =   1680
      TabIndex        =   14
      Top             =   1440
      Width           =   2535
   End
   Begin VB.CommandButton Command5 
      Caption         =   "F2                                                                 Observações                                              "
      Height          =   975
      Left            =   14640
      TabIndex        =   13
      Top             =   8520
      Width           =   2535
   End
   Begin VB.CommandButton Command4 
      Caption         =   "F5                                                                    Acréssimo                                              "
      Height          =   975
      Left            =   15240
      TabIndex        =   12
      Top             =   8160
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "F1                                                                      Ajuda                                              "
      Height          =   855
      Left            =   14760
      TabIndex        =   11
      Top             =   8520
      Width           =   2535
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      MaxLength       =   65
      TabIndex        =   10
      Top             =   840
      Width           =   6015
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8640
      TabIndex        =   9
      Top             =   120
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   8
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "F4                                                                     Tamanho                                              "
      Height          =   855
      Left            =   15000
      TabIndex        =   6
      Top             =   8160
      Width           =   2535
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   6735
      Left            =   4440
      Negotiate       =   -1  'True
      TabIndex        =   3
      Top             =   2760
      Visible         =   0   'False
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   11880
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   24
      WrapCellPointer =   -1  'True
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
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Selecione o produto "
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
   Begin VB.CommandButton Command1 
      Caption         =   "F3                                                                 Ingredientes                                             "
      Height          =   855
      Left            =   14760
      TabIndex        =   2
      Top             =   8280
      Width           =   2535
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   9585
      Width           =   19080
      _ExtentX        =   33655
      _ExtentY        =   873
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.TreeView TreeView1 
      Height          =   9375
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   16536
      _Version        =   327682
      Style           =   7
      Appearance      =   1
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
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   7455
      Left            =   17040
      TabIndex        =   19
      Top             =   480
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   13150
      _Version        =   393216
      BackColor       =   12632319
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
      Caption         =   "Pedido "
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
   Begin VB.Label Label5 
      Caption         =   "R$"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12000
      TabIndex        =   22
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "Codigo do sistema"
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
      Left            =   7560
      TabIndex        =   7
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Produto descrição"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   5
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Meu Codigo"
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
      Left            =   4560
      TabIndex        =   4
      Top             =   240
      Width           =   1095
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   3840
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   11
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "NewatendimentoVendas.frx":0A02
            Key             =   "mais"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "NewatendimentoVendas.frx":0BDC
            Key             =   "check"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "NewatendimentoVendas.frx":0DB6
            Key             =   "pasta"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "NewatendimentoVendas.frx":0F90
            Key             =   "filtrar"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "NewatendimentoVendas.frx":116A
            Key             =   "localizar"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "NewatendimentoVendas.frx":1344
            Key             =   "voltar"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "NewatendimentoVendas.frx":151E
            Key             =   "partir"
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "NewatendimentoVendas.frx":16F8
            Key             =   "voltar2"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "NewatendimentoVendas.frx":18D2
            Key             =   "mala"
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "NewatendimentoVendas.frx":1AAC
            Key             =   "voltaraocomerco"
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "NewatendimentoVendas.frx":1C86
            Key             =   "Carteira"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnu_atendimento 
      Caption         =   "Atendimento"
   End
End
Attribute VB_Name = "Form440"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim mcODIDGO As Integer
Dim CODIGOsIS As Integer
Dim DESCRITION As String
Dim mobilidadeMedida As Integer
Dim metadedeProtuto  As Integer
Dim contadorDeMetade As Integer

Dim ValordeGrid As String

Dim i As Integer



Public Sub criaritens()
Dim totalDecardapiosListados As Integer
Dim coutCardapioTitulo As Integer
Dim procurarSub As String

ConServer


Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs.ActiveConnection = con

rs.CursorLocation = adUseClient
  sql = "SELECT COUNT(DISTINCT `Titulo`) FROM `Cardapio` WHERE 1"
  rs.Open sql
  coutCardapioTitulo = rs.Fields("COUNT(DISTINCT `Titulo`)").Value

Dim nodx As Node

'limpa qualquer nó criado
TreeView1.Nodes.Clear

Set TreeView1.ImageList = ImageList1

'novabuscapelos titulos
  rs.Close
  sql = "SELECT DISTINCT`Titulo` FROM `Cardapio` WHERE 1"
  rs.Open sql
 'coutCardapioTitulo = rs.Fields("COUNT(DISTINCT `Titulo`)").Value

For i = coutCardapioTitulo To 1 Step -1
 

'Adicionar titulo
 Dim sKey As String
    Dim oNodex As Node
    
''Adicione alguns itens de nível raiz
''
'para i = 1 a 5
'TreeView1.Nodes.Add , , "ROOT" & i, "Item Raiz" & i
TreeView1.Nodes.Add , , "ROOT" & i, rs.Fields("Titulo").Value
totalDecardapiosListados = totalDecardapiosListados + 1
rs.MoveNext


Next

'próximo
'"
'Agora adicionar alguns filhos
'"
'para i = 1 a 5
'com TreeView1.Nodes
'.Adicionar "ROOT1", tvwChild, "ROOT1CHILD" & i, "Item infantil" & i
'.Adicionar "ROOT2", tvwChild, "ROOT2CHILD" & i, "Item infantil" & i
'.Adicionar "ROOT3", tvwChild, "ROOT3CHILD" & i, "Item para crianças" & i
'.Adicione "ROOT4", tvwChild, "ROOT4CHILD" & i, "Item infantil" & i
'.Adicione "ROOT5", tvwChild, "ROOT5CHILD" & i, "Item para crianças" & i
'End With
'Next
''
' definir  o quantos subs tem para serem creidos
Dim t As Integer
Dim redusCardapios As Integer
redusCardapios = totalDecardapiosListados
For t = 1 To totalDecardapiosListados
                 
                procurarSub = TreeView1.Nodes.Item(t)
                rs.Close
                  sql = "SELECT DISTINCT `codTitulo` FROM `Cardapio` WHERE `Titulo` LIKE '" & procurarSub & "' "
                  rs.Open sql
                  Dim retornaCodigodoiten As Integer
                  
                  retornaCodigodoiten = rs.Fields("codTitulo").Value
                 rs.Close
                  sql = "SELECT COUNT(DISTINCT `tipo`) FROM `Cardapio` WHERE `codTitulo`='" & retornaCodigodoiten & "'"
                  rs.Open sql
                  coutCardapioTitulo = rs.Fields("COUNT(DISTINCT `tipo`)").Value
                
                
                rs.Close
                 sql = "SELECT DISTINCT `tipo` FROM `Cardapio` WHERE `codTitulo`= '" & retornaCodigodoiten & "'"
                rs.Open sql
                For i = 1 To coutCardapioTitulo
                With TreeView1.Nodes
                
                If (rs.Fields("tipo").Value <> "") Then
                .Add "ROOT" & redusCardapios, tvwChild, "ROOT" & redusCardapios & "CHILD" & i, rs.Fields("tipo").Value
                 End If
                End With
                rs.MoveNext
                Next
redusCardapios = redusCardapios - 1
Next

'' Agora adicione alguns Grand- Children
''
'para i = 1 a 5
'com TreeView1.Nodes
'.Add "ROOT1CHILD2", tvwChild, "Grand criança" & i
'.Add "ROOT2CHILD2", tvwChild, "Grand criança" & i
'.Add "ROOT3CHILD2", tvwChild, , "Grand filho" & i
'.Adicionar "ROOT4CHILD2", tvwChild, "Grand filho" & i
'.Adicione "ROOT5CHILD2", tvwChild, "Grand Child" e termino
'com o
'próximo
'For i = 1 To 5
'With TreeView1.Nodes
'.Add "ROOT1CHILD2", tvwChild, , "Grand Child " & i
'.Add "ROOT2CHILD2", tvwChild, , "Grand Child " & i
'.Add "ROOT3CHILD2", tvwChild, , "Grand Child " & i
'.Add "ROOT4CHILD2", tvwChild, , "Grand Child " & i
'.Add "ROOT5CHILD2", tvwChild, , "Grand Child " & i
'End With
'Next














 Set rs = Nothing

     
End Sub


Private Sub Command6_Click()
metadedeum
If Text3.Text <> "" Then
definaAsmedidaspossiveis (Text3.Text)
End If

End Sub

Private Sub DataGrid1_Click()
RetornoPordescricao (DataGrid1.Columns(0).Value)
End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF6 Then
MsgBox "1/2 "
            metadedeProtuto = 1
                If Text3.Text = DataGrid1.Columns(0).Value Then
            
            'Else
                    If (DataGrid1.Columns(0).Value = "") Then
                    RetornoPordescricao (ValordeGrid)
                    Else
                    RetornoPordescricao (DataGrid1.Columns(0).Value)
                    End If
            End If
                
  End If
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 47 Then
MsgBox "1/2 "
metadedeProtuto = 1
End If

If KeyAscii = 13 Then
 KeyAscii = 0
                   If mobilidadeMedida = 1 Then
                   Text2.Text = DataGrid1.Columns(1).Value
                       If Text2.Text = DataGrid1.Columns(1).Value Then
                            MsgBox "MEDIDA  FEITA !   "
                            definiroPreco (Text2.Text)
                            'definaAsmedidaspossiveis (Text3.Text)
                            mobilidadeMedida = 0
                        End If
       'MOSTRAR O PREÇO
       If metadedeProtuto = 1 Then
               
      Else
      
      
      End If
      
         
         
                  Else
                                If Text3.Text = DataGrid1.Columns(0).Value Then
                                   MsgBox "ESCOLHA FEITA !   "
                                   If metadedeProtuto <> 1 Then
                                   definaAsmedidaspossiveis (Text3.Text)
                                   Else
                                   metadedeum
                                   Text3.SetFocus
                                   
                   End If
                   
                   Else
                           If (DataGrid1.Columns(0).Value = "") Then
                           RetornoPordescricao (ValordeGrid)
                           Else
                           RetornoPordescricao (DataGrid1.Columns(0).Value)
                           End If
                   End If
          End If
End If
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
ValordeGrid = DataGrid1.Columns(0).Value

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)





If KeyCode = vbKeyF6 Then

                If Text3.Text = DataGrid1.Columns(0).Value Then
            
            'Else
                    If (DataGrid1.Columns(0).Value = "") Then
                    RetornoPordescricao (ValordeGrid)
                    Else
                    RetornoPordescricao (DataGrid1.Columns(0).Value)
                    End If
            End If
                
  End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

'voltar ao escolhar

If KeyAscii = vbKeyEscape Then
CardapioPordescricao (Text3.Text)
mobilidadeMedida = 0
Text7.Text = ""
Text7.Visible = False
End If
If KeyAscii = 47 Then
KeyAscii = 0
MsgBox "1/2 "
metadedeProtuto = 1
End If






End Sub

Private Sub Form_Load()
criaritens
End Sub

Private Sub Text1_Change()
mobilidadeMedida = 0
End Sub

Private Sub Text1_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text9.Text = ""

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = (SoNumeros(KeyAscii))
 If KeyAscii = 13 Then
   KeyAscii = 0
CardapioPorMeuCodigo (Text1.Text)
'Text2.Text = ""
'Text3.Text = ""

End If
 If KeyAscii = 8 Then
   KeyAscii = 0
   Text9.Text = ""
'Text1.Text = ""
'Text2.Text = ""
'Text3.Text = ""
End If
End Sub

Private Sub Text2_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text9.Text = ""
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
KeyAscii = (SoNumeros(KeyAscii))
 If KeyAscii = 13 Then
   KeyAscii = 0
CardapioCodigosistem (Text2.Text)
'Text1.Text = ""
'Text3.Text = ""

End If
 If KeyAscii = 8 Then
   KeyAscii = 0
   Text9.Text = ""
'Text2.Text = ""
'Text1.Text = ""
'Text3.Text = ""
End If
End Sub

Private Sub Text3_Change()
mobilidadeMedida = 0
End Sub

Private Sub Text3_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text9.Text = ""

End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
   KeyAscii = 0
CardapioPordescricao (Text3.Text)
'Text1.Text = ""
'Text2.Text = ""
End If

 If KeyAscii = 8 Then
   KeyAscii = 0
Text9.Text = ""
Text3.Text = ""
End If

End Sub

Private Sub TreeView1_NodeClick(ByVal Node As ComctlLib.Node)
Dim Proc As String

Proc = Trim("%" + TreeView1.SelectedItem + "%")
ConServer


Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs.ActiveConnection = con
rs.CursorLocation = adUseClient


  
         



  sql = "SELECT DISTINCT `Descricao` FROM `Cardapio` WHERE `tipo` LIKE '" & Proc & "' ORDER BY `Cardapio`.`Descricao` ASC"
  rs.Open sql
 If rs.BOF = False Then
 Set DataGrid1.DataSource = rs
 DataGrid1.Visible = True
 DataGrid1.SetFocus
 Else
 DataGrid1.Visible = False
 
 End If



 Set rs = Nothing
DataGrid1.Columns(0).Caption = TreeView1.SelectedItem
DataGrid1.Columns(0).Width = 12000
End Sub


Public Sub CardapioPorMeuCodigo(Mcodigo As Integer)
Dim intMcodigo As Integer
intMcodigo = Mcodigo

ConServer

Dim NpTitulo As String

Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs.ActiveConnection = con
rs.CursorLocation = adUseClient


'RETORNOO DE VARIAVAIS
                                         sql = "SELECT `idCardapio`,`codigo`,`Descricao` ,`tipo`FROM `Cardapio` WHERE `codigo` = '" & intMcodigo & "'ORDER BY `Cardapio`.`Descricao` ASC "
                                                rs.Open sql
                                                 If rs.BOF = False Then
                                                  mcODIDGO = rs.Fields("codigo").Value
                                                  CODIGOsIS = rs.Fields("idCardapio").Value
                                                  DESCRITION = rs.Fields("Descricao").Value
                                                  NpTitulo = rs.Fields("tipo").Value
                                                End If
                                                rs.Close
                                                 
                                        

        
         
  
  sql = "SELECT `Descricao` FROM `Cardapio` WHERE `codigo` = '" & intMcodigo & "'ORDER BY `Cardapio`.`Descricao` ASC "
  rs.Open sql
 If rs.BOF = False Then
 Set DataGrid1.DataSource = rs
 DataGrid1.Visible = True
 DataGrid1.SetFocus
 Else
 MsgBox " O codigo : '" & Mcodigo & "'não foi localizado no sistema ", , "Não Existe"
 DataGrid1.Visible = False
 
 End If


'rs.Close
                                                  'mcODIDGO = rs.Fields("codigo").Value
                                                  Text2.Text = CODIGOsIS
                                                  Text3.Text = DESCRITION
                                                  
DataGrid1.Columns(0).Caption = NpTitulo


DataGrid1.Columns(0).Width = 12000

 Set rs = Nothing
End Sub

Public Sub CardapioCodigosistem(Mcodigo As Integer)
Dim intMcodigo As Integer
intMcodigo = Mcodigo

ConServer

Dim NpTitulo As String

Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs.ActiveConnection = con
rs.CursorLocation = adUseClient

       ' sql = "SELECT `tipo` FROM `Cardapio` WHERE `idCardapio` = '" & intMcodigo & "'ORDER BY `Cardapio`.`Descricao` ASC "
      'RETORNOO DE VARIAVAIS
                                         sql = "SELECT `idCardapio`,`codigo`,`Descricao` ,`tipo`FROM `Cardapio` WHERE `idCardapio` = '" & intMcodigo & "'ORDER BY `Cardapio`.`Descricao` ASC "
                                                rs.Open sql
                                                 If rs.BOF = False Then
                                                  mcODIDGO = rs.Fields("codigo").Value
                                                  CODIGOsIS = rs.Fields("idCardapio").Value
                                                  DESCRITION = rs.Fields("Descricao").Value
                                                  NpTitulo = rs.Fields("tipo").Value
                                                End If
                                                rs.Close
         
  
  sql = "SELECT `Descricao` FROM `Cardapio` WHERE `idCardapio` = '" & intMcodigo & "'ORDER BY `Cardapio`.`Descricao` ASC "
  rs.Open sql
 If rs.BOF = False Then
 Set DataGrid1.DataSource = rs
 DataGrid1.Visible = True
 DataGrid1.SetFocus
 Else
 MsgBox " O codigo : '" & Mcodigo & "'não foi localizado no sistema ", , "Não Existe"
 DataGrid1.Visible = False
 
 End If


'rs.Close
                                                  Text1.Text = mcODIDGO
                                                  'Text2.Text = CODIGOsIS
                                                  Text3.Text = DESCRITION
DataGrid1.Columns(0).Caption = NpTitulo
DataGrid1.Columns(0).Width = 12000

 Set rs = Nothing
End Sub

Public Sub CardapioPordescricao(Mcodigo As String)
Dim intMcodigo As String
intMcodigo = Trim("%" + Mcodigo + "%")

ConServer

Dim NpTitulo As String

Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs.ActiveConnection = con
rs.CursorLocation = adUseClient
        

          'sql = "SELECT `tipo` FROM `Cardapio` WHERE  `Descricao` LIKE'" & intMcodigo & "'ORDER BY `Cardapio`.`Descricao` ASC "
       'RETORNOO DE VARIAVAIS
                                         sql = "SELECT `idCardapio`,`codigo`,`Descricao`,`tipo` FROM `Cardapio` WHERE `Descricao` LIKE'" & intMcodigo & "'ORDER BY `Cardapio`.`Descricao` ASC "
                                                rs.Open sql
                                                 If rs.BOF = False Then
                                                  mcODIDGO = rs.Fields("codigo").Value
                                                  CODIGOsIS = rs.Fields("idCardapio").Value
                                                  DESCRITION = rs.Fields("Descricao").Value
                                                  NpTitulo = rs.Fields("tipo").Value
                                                End If
                                                rs.Close
                 
  
  sql = "SELECT DISTINCT `Descricao` FROM `Cardapio` WHERE `Descricao` LIKE'" & intMcodigo & "'ORDER BY `Cardapio`.`Descricao` ASC "
  rs.Open sql
 If rs.BOF = False Then
 Set DataGrid1.DataSource = rs
 DataGrid1.Visible = True
 DataGrid1.SetFocus
 Else
 MsgBox " O produto : '" & Mcodigo & "'não foi localizado no sistema ", , "Não Existe"
 DataGrid1.Visible = False
 
 End If


'rs.Close
                                                  Text1.Text = mcODIDGO
                                                  Text2.Text = CODIGOsIS
                                                  'Text3.Text = DESCRITION
DataGrid1.Columns(0).Caption = NpTitulo
DataGrid1.Columns(0).Width = 12000

 Set rs = Nothing
End Sub

Public Sub RetornoPordescricao(Mcodigo As String)

Dim intMcodigo As String
intMcodigo = Mcodigo
If (intMcodigo <> "") Then
ConServer

Dim NpTitulo As String

Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs.ActiveConnection = con
rs.CursorLocation = adUseClient
        

          'sql = "SELECT `tipo` FROM `Cardapio` WHERE  `Descricao` LIKE'" & intMcodigo & "'ORDER BY `Cardapio`.`Descricao` ASC "
       'RETORNOO DE VARIAVAIS
                                         sql = "SELECT `idCardapio`,`codigo`,`Descricao`,`tipo` FROM `Cardapio` WHERE `Descricao` LIKE'" & intMcodigo & "'ORDER BY `Cardapio`.`Descricao` ASC "
                                                rs.Open sql
                                                 If rs.BOF = False Then
                                                  mcODIDGO = rs.Fields("codigo").Value
                                                  CODIGOsIS = rs.Fields("idCardapio").Value
                                                  DESCRITION = rs.Fields("Descricao").Value
                                                  NpTitulo = rs.Fields("tipo").Value
                                                End If
                                                rs.Close
                 
  
 
 
                                                   Text1.Text = mcODIDGO
                                                  Text2.Text = CODIGOsIS
                                                  Text3.Text = DESCRITION
DataGrid1.Columns(0).Caption = NpTitulo
DataGrid1.Columns(0).Width = 12000
End If
 Set rs = Nothing
End Sub


Public Sub definaAsmedidaspossiveis(Mcodigo As String)
Dim UNidadeDisponiveis As Integer
Dim intMcodigo As String
intMcodigo = Trim("%" + Mcodigo + "%")

ConServer

Dim NpTitulo As String

Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs.ActiveConnection = con
rs.CursorLocation = adUseClient
        

          '
       'RETORNOO DE VARIAVAIS
                                        '  sql = "SELECT COUNT (`Medida`) FROM `Cardapio` WHERE `Descricao` LIKE "
                                                sql = "SELECT COUNT(`Medida`) FROM `Cardapio` WHERE `Descricao` LIKE '" & intMcodigo & "'ORDER BY `Medida` ASC"
                                                
                                                rs.Open sql
                                                UNidadeDisponiveis = rs.Fields("COUNT(`Medida`)").Value

                                                rs.Close
       
                                         sql = "SELECT `Medida`,`idCardapio` FROM `Cardapio` WHERE `Descricao` LIKE '" & intMcodigo & "' ORDER BY `Cardapio`.`Medida` ASC"
                                                rs.Open sql
                                                 If rs.BOF = False And UNidadeDisponiveis > 1 Then
                                                  'mcODIDGO = rs.Fields("codigo").Value
                                                  Set DataGrid1.DataSource = rs
                                                  DataGrid1.Caption = "Escolha entre as medidas disponíveis"
                                                  DataGrid1.Visible = True
                                                  DataGrid1.SetFocus
                                                  mobilidadeMedida = 1
                                               Else
                 
  

 MsgBox " O produto  : '" & Mcodigo & "' sem medidas disponiveis  ", , "Não existe medida para este produto"
 DataGrid1.Visible = False
 
 End If


DataGrid1.Columns(0).Caption = "MEDIDAS POSSÍVEIS "
DataGrid1.Columns(0).Width = 12000

 Set rs = Nothing

End Sub
Public Sub definiroPreco(Mcodigo As Integer)
Dim intMcodigo As Integer
intMcodigo = Mcodigo

ConServer

Dim NpTitulo As String

Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs.ActiveConnection = con
rs.CursorLocation = adUseClient

      
                                         sql = "SELECT `valor` ,`Medida` FROM `Cardapio` WHERE `idCardapio` = '" & intMcodigo & "'ORDER BY `Cardapio`.`Descricao` ASC "
                                                rs.Open sql
                                                 If rs.BOF = False Then
                                                 
                                                 Text7.Visible = False
                                                  Text9.Text = rs.Fields("valor").Value
                                                  End If
                                                
                                                
                                                rs.Close
         
  
 
    
 DataGrid1.Visible = False
 


 Set rs = Nothing
End Sub

Public Sub metadedeum()
'pega o texto3 transfere para o 7 adicionado o sinal
                                     If metadedeProtuto = 1 Then
                                                     Text7.Visible = True
                                                            If Text7 = "" Then
                                                            'definir o tamanho antes
                                                            definiroPreco (Text3.Text)
                                                             Text9.Text = Text9 / 2
                                                            contadorDeMetade = 1
                                                            Text7.Text = "1/2 " & Text3.Text
                                                            Text3.Text = ""
                                                            Else
                                                            contadorDeMetade = 2
                                                            Text7.Text = Text7.Text & " _  1/2  " & Text3.Text
                                                            End If
                                               If contadorDeMetade = 2 Then
                                                'Text7.Visible = False
                                                 metadedeProtuto = 0
                                              End If
                                             End If





'divide o valor

End Sub
