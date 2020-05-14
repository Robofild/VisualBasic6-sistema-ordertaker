VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form Form407 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pagamento"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8055
   Icon            =   "PagamentoTotalies.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form15"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   8055
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   13680
      TabIndex        =   28
      Text            =   "Text8"
      Top             =   1320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   13560
      TabIndex        =   25
      Text            =   "Text7"
      Top             =   600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   11280
      TabIndex        =   22
      Text            =   "Text6"
      Top             =   1560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   11400
      TabIndex        =   19
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Remover"
      Height          =   615
      Left            =   7080
      Picture         =   "PagamentoTotalies.frx":0A02
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1440
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "reinforma"
      Height          =   495
      Left            =   10200
      TabIndex        =   14
      Top             =   2040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   10080
      TabIndex        =   13
      Top             =   2760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3495
      Left            =   240
      TabIndex        =   10
      Top             =   1200
      Visible         =   0   'False
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   6165
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
      Caption         =   "Pagamentos "
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
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   10200
      TabIndex        =   9
      Text            =   "Text4"
      Top             =   3840
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   8760
      TabIndex        =   8
      Text            =   "Text3"
      Top             =   3840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   7200
      TabIndex        =   7
      Text            =   "Text2"
      Top             =   3840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "F1 Formas de Pagamento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5280
      Picture         =   "PagamentoTotalies.frx":1404
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Text            =   "71,50"
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label12 
      Caption         =   "retorno txt2 406"
      Height          =   375
      Left            =   13560
      TabIndex        =   29
      Top             =   960
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label11 
      Caption         =   "tipo de cupom"
      Height          =   255
      Left            =   14760
      TabIndex        =   27
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label10 
      Caption         =   "Label10"
      Height          =   375
      Left            =   14760
      TabIndex        =   26
      Top             =   480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "frete anterior"
      Height          =   255
      Left            =   13200
      TabIndex        =   24
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
      Height          =   375
      Left            =   1440
      TabIndex        =   23
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "Valor Resolvido"
      Height          =   255
      Left            =   3600
      TabIndex        =   21
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Falta :"
      Height          =   195
      Left            =   3600
      TabIndex        =   20
      Top             =   840
      Width           =   480
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   255
      Left            =   4920
      TabIndex        =   17
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "label4"
      Height          =   255
      Left            =   4200
      TabIndex        =   16
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   360
      Left            =   3720
      TabIndex        =   15
      Top             =   360
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   375
      Left            =   10320
      TabIndex        =   12
      Top             =   840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "226"
      Height          =   375
      Left            =   10080
      TabIndex        =   11
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label21 
      Caption         =   "0"
      Height          =   495
      Left            =   7920
      TabIndex        =   6
      Top             =   2760
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label20 
      Caption         =   "Label20"
      Height          =   375
      Left            =   8040
      TabIndex        =   5
      Top             =   2040
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label19 
      Caption         =   "fP"
      Height          =   375
      Left            =   8640
      TabIndex        =   4
      Top             =   1440
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label18 
      Caption         =   "Label18"
      Height          =   375
      Left            =   8400
      TabIndex        =   3
      Top             =   840
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label17 
      Caption         =   "ENTREGA"
      Height          =   255
      Left            =   7920
      TabIndex        =   2
      Top             =   360
      Visible         =   0   'False
      Width           =   1455
   End
End
Attribute VB_Name = "Form407"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Cancelar As Boolean
Dim valor1 As Double
Dim valor2 As Double
Dim valor3 As Double
Dim Revalidar As Boolean
Dim controleTrasito As Boolean
Dim escolhafeita As Boolean

Private Sub Command1_Click()

'Dim pr as New prin
If Text6.Text = "1" Then
rEATIVARfUNCOESantigas
Else
Form408.Show
Form409.Label2.Caption = "Dinheiro"
Form409.Label7.Caption = Replace(Text1.Text, ",", ".")

End If
End Sub

Private Sub Command2_Click()
ConServer

Dim somaValorPago As Double

Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset
Set rs.ActiveConnection = con


rs.CursorLocation = adUseClient

 'apanhar numero do  PEDIDO OFICIAL
        
         sql = "SELECT`FormaPagamento`,`ReceberPago`,`ValorRecebido`,`ValorPago`,`Troco`,`NUmPedido` ,`id`FROM `Pagamento` WHERE `NUmPedido`='" & Label1.Caption & "' ORDER BY `Pagamento`.`FormaPagamento` DESC"
         rs.Open sql
          If rs.BOF = False Then
          Set DataGrid1.DataSource = rs
          DataGrid1.Columns(0).Caption = "Forma de Pagamento"
          DataGrid1.Columns(1).Caption = "Situação de Pagamento"
          DataGrid1.Columns(2).Caption = "Valor Rec"
          DataGrid1.Columns(3).Caption = "Valor PG"
          DataGrid1.Columns(4).Caption = "Troco"
          DataGrid1.Columns(5).Caption = "Nº Pedido"
         
           DataGrid1.Visible = True
           Command4.Visible = True
           abilitelabeis
           'Form407.Height = 5160
          Else
            'Form407.Height = 1545
          DataGrid1.Visible = False
          Command4.Visible = False
          falsebilitelabeis
          Text6.Text = 0
          Command1.Caption = "F1 Formas de Pagamento"
          End If
          
 Set rs = Nothing


End Sub

Private Sub Command3_Click()
ConServer

Dim somaValorPago As Double

Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset
Set rs.ActiveConnection = con


rs.CursorLocation = adUseClient

 'apanhar numero do  PEDIDO OFICIAL
        
         sql = "SELECT SUM(`ValorPago`) FROM `Pagamento` WHERE `NUmPedido`='" & Label1.Caption & "'"
         'sql = "SELECT SUM(`ValorPago`) FROM `Pagamento` WHERE `NUmPedido`='226'"
         rs.Open sql
          If rs.BOF = False Then
          If IsNull(rs.Fields("SUM(`ValorPago`)")) Then
          Else
           somaValorPago = Format(rs.Fields("SUM(`ValorPago`)").Value, "General Number")
           Label3.Caption = "( " & somaValorPago & " )"
           Label5 = somaValorPago
           Label4 = Text1 - Label5
           Label3.Visible = True
           End If
          Else
          Label3.Visible = False
          End If
          
 Set rs = Nothing

If Label4 = "0" Then
niveldeFechamento
Text6.Text = 1
Else
Command1.Caption = "F1 Formas de Pagamento"

Command1.BackColor = &H8000000F
Text6.Text = 0

End If

End Sub

Public Sub AbertoParaPagar()
Command1.Caption = "F1 Formas de Pagamento"
Command1.BackColor = &H8000000F
Text6.Text = 0
End Sub
Public Sub niveldeFechamento()
Command1.Caption = "F1 Finalizar o pedido!"
Command1.BackColor = &HFF00&
Text6.Text = 1
End Sub

Private Sub Command4_Click()

ConServer

Dim somaValorPago As Double
Dim remId As Integer
Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset
Set rs.ActiveConnection = con


rs.CursorLocation = adUseClient
remId = DataGrid1.Columns(6).Value
 'apanhar numero do  PEDIDO OFICIAL
        
         sql = "DELETE FROM `Pagamento` WHERE `id`='" & remId & "'"
         rs.Open sql
          MsgBox "pagamento excluido com sucesso !", , "Pagamento Excluido"
        Call Command2_Click
        Call Command3_Click
'        rs.Close
             sql = "SELECT`FormaPagamento`,`ReceberPago`,`ValorRecebido`,`ValorPago`,`Troco`,`NUmPedido` ,`id`FROM `Pagamento` WHERE `NUmPedido`='" & Label1.Caption & "' ORDER BY `Pagamento`.`FormaPagamento` DESC"
         rs.Open sql
          If rs.BOF = True Then
          Unload Me
          Form406.Show
        End If
        
        
        
        
 Set rs = Nothing
If Label4 <> "0" Then
AbertoParaPagar
Else
niveldeFechamento
End If
End Sub


Private Sub DataGrid1_DblClick()
Dim valordoLabel4 As Single
valordoLabel4 = Label4


RecondicionarValorDePagamento (True)

End Sub

Private Sub Form_Load()
 Form406.Hide

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF1 Then
   Call Command1_Click
  End If

 
    
 
End Sub
Public Sub activid()
       Call Command2_Click
        Call Command3_Click
        Text5.Text = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If Text7.Text <> "0" And Text7.Text <> "Text7" Then
ConServer

Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs.ActiveConnection = con
rs.CursorLocation = adUseClient
sql = "SELECT * FROM `at_frete` WHERE `fk_numPedido` =  '" & Label1.Caption & "' ORDER BY `fk_numPedido` DESC"
rs.Open sql
If rs.BOF = True Then
rs.Close
sql = "INSERT INTO `at_frete` (`id`, `frete`, `fk_numPedido`) VALUES (NULL, '" & Text7.Text & "',  '" & Label1.Caption & "') "
                            rs.Open sql
End If

 Set rs = Nothing

End If
End Sub

Private Sub Label1_Change()
activid
End Sub

Private Sub Label4_Change()
Dim valordoLabel4 As Single
If Label4.Caption = "0" Then
MsgBox "F1 Pagamento concluido com sucesso", , "Finalize a operação (F1)"""
Text6.Text = 1
Else
If Label4 <> "label4" Then
valordoLabel4 = Label4
End If

End If
If valordoLabel4 < 0 Then

End If
End Sub

Private Sub Text1_Change()
Dim valorparasoma As Single
Label8.Caption = DBLcurre(Text1.Text)
valorparasoma = Replace(Label8.Caption, ".", ",")
If Label3.Caption <> "Label3" Then
Label4 = -Label3
End If
If Label4 <> "0" Then
If Label4 <> "label4" Then
AbertoParaPagar
Text6.Text = 0
End If
Else
If Label4 <> "label4" Then
niveldeFechamento
Text6.Text = 1
End If
End If
Call Command3_Click

End Sub

Private Sub Text5_Change()
If Text5.Text = 1 Then

activid

End If
End Sub

Public Sub falsebilitelabeis()
Label3.Visible = False
Label5.Visible = False
Label4.Visible = False
Label7.Visible = False
Label6.Visible = False

End Sub
Public Sub abilitelabeis()
Label3.Visible = True
Label5.Visible = False
Label4.Visible = True
Label7.Visible = True
Label6.Visible = True

End Sub

Public Sub rEATIVARfUNCOESantigas()
escolhafeita = True
Cancelar = True
Form501.Show
Form503.Show
Form501.Timer1.Interval = 100
End Sub


Public Function RecondicionarValorDePagamento(recondicionar As Boolean) As Boolean
'MsgBox "Clique sobre o pagamento para alterar", vbInformation, "Altere o pagamento!"
Form409.Text4.Text = 1
Form409.Text1.Text = DataGrid1.Columns(2).Value
Form409.Text2.Text = DataGrid1.Columns(3).Value - (Label4 * (-1))
Form409.Text3.Text = DataGrid1.Columns(4).Value
Form409.Label2.Caption = DataGrid1.Columns(0).Value
Form409.Text5.Text = DataGrid1.Columns(6).Value
Form409.Caption = "Recondicionado o pagamento pedido" & Label1.Caption
Form409.Show

End Function

Private Sub Text8_Change()
Form409.Text2.Text = Text8

End Sub
