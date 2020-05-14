VERSION 5.00
Begin VB.Form Form409 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Valor  a  decidir "
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6945
   Icon            =   "PagamentoDefine.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form16"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   6945
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   7440
      TabIndex        =   20
      Text            =   "Text6"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Caption         =   "F2"
      Height          =   615
      Left            =   1320
      TabIndex        =   18
      Top             =   1920
      Width           =   1455
      Begin VB.CheckBox Check1 
         Caption         =   "Pago"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   10440
      TabIndex        =   13
      Text            =   "Text5"
      Top             =   600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   10440
      TabIndex        =   11
      Text            =   "0"
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "F1-  Confirmar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1680
      Picture         =   "PagamentoDefine.frx":0A02
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2160
      Width           =   3975
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4920
      TabIndex        =   2
      Text            =   "0"
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2640
      TabIndex        =   1
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label12 
      Caption         =   "Troco"
      Height          =   255
      Left            =   5040
      TabIndex        =   17
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label11 
      Caption         =   "Valor a pagar "
      Height          =   255
      Left            =   2640
      TabIndex        =   16
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label10 
      Caption         =   "Troco para:?"
      Height          =   255
      Left            =   360
      TabIndex        =   15
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label9 
      Caption         =   "indice"
      Height          =   375
      Left            =   9960
      TabIndex        =   14
      Top             =   600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label8 
      Caption         =   "recondi"
      Height          =   255
      Left            =   9720
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   375
      Left            =   8280
      TabIndex        =   10
      Top             =   3000
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   375
      Left            =   8400
      TabIndex        =   9
      Top             =   2400
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   375
      Left            =   8400
      TabIndex        =   8
      Top             =   1800
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   375
      Left            =   8400
      TabIndex        =   7
      Top             =   1200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   375
      Left            =   8400
      TabIndex        =   6
      Top             =   720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   8400
      TabIndex        =   5
      Top             =   360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Tipo de pag"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   3
      Top             =   0
      Width           =   6255
   End
End
Attribute VB_Name = "Form409"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim valorDeBITO

Private Sub Command1_Click()
If Text4.Text = "1" Then
ConServer


Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs.ActiveConnection = con

Dim valorDosItens As Double
        If Check1.Value = 1 Then
        Label5.Caption = "Pago"
        Else
        Label5.Caption = "Receber"
        
        End If
        
Observacoess (Trim(Label2.Caption))
  sql = "UPDATE `robofi61_order_taker`.`Pagamento` SET `EntregarBuscar` = ' " & Form407.Label17.Caption & "', `ReceberPago` = ' " & Format(Form409.Label5.Caption, ">") & "', `FormaPagamento` = ' " & Trim(Format(Form409.Label2.Caption, ">")) & "', `ValorTotal` = ' " & Form407.Label8.Caption & "', `ValorRecebido` = ' " & Form409.Label1.Caption & "', `ValorPago` = ' " & Form409.Label3.Caption & "', `Troco` = ' " & Form409.Label4.Caption & "', `NUmPedido` = ' " & Form407.Label1.Caption & "',`NumClient` = ' " & Form2.Text11.Text & "', `Observação` = ' " & Form409.Label6.Caption & "' WHERE `id`='" & Text5.Text & "' "
'UPDATE `robofi61_order_taker`.`Pagamento` SET `EntregarBuscar` = ' ENTREGAs', `ReceberPago` = ' Receber s', `FormaPagamento` = ' Dinheiros ', `ValorTotal` = '15', `ValorRecebido` = '50', `ValorPago` = '15', `Troco` = '35', `NUmPedido` = '2311', `NumClient` = '10', `Observação` = ' # NAOs LEVAR TROCO  ' WHERE ;

  rs.Open sql
  Set rs = Nothing
calcule
If Label4.Caption = "0" Then
Text4.Text = "1"
End If

Else

        calcule
        If Check1.Value = 1 Then
        Label5.Caption = "Pago"
        Else
        Label5.Caption = "Receber"
        
        End If
End If
Observacoess (Trim(Label2.Caption))

Unload Me
Unload Form408
Unload Form406
Form407.Text5.Text = 1
Form407.Show

End Sub

Public Sub calcule()
If Text1 <> "" And Text2 <> "" Then
Text3 = Text1 - Text2
End If
End Sub

Private Sub Command1_GotFocus()
If Text1.Text = "" Then
Text1.Text = Text2.Text
End If
Label1.Caption = Replace(Text1.Text, ",", ".")
Label3.Caption = Replace(Text2.Text, ",", ".")
Label4.Caption = Replace(Text3.Text, ",", ".")

End Sub

Private Sub Form_Load()
If Form407.Label3.Visible = True Then
  valorDeBITO = Form407.Label4.Caption
Else
    valorDeBITO = Form407.Text1
End If
Text2.Text = valorDeBITO
Label7.Caption = Replace(Form407.Text1.Text, ",", ".")
End Sub

Private Sub Text1_Change()
Label1.Caption = Replace(Text1.Text, ",", ".")
calcule
End Sub

Private Sub Text1_GotFocus()
Text1.Text = ""
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF1 Then
  Call Command1_Click
  End If
    If KeyCode = vbKeyF2 Then
    If Check1.Value = 0 Then
   Check1.Value = 1
   Else
   Check1.Value = 0
   End If
  End If

End Sub
Private Sub Text1_LostFocus()



calcule
End Sub

Private Sub Text2_Change()

Label3.Caption = Replace(Text2.Text, ",", ".")
Label4.Caption = Replace(Text3.Text, ",", ".")




End Sub

Public Sub Observacoess(pagamento As String)
Select Case pagamento
Case "Dinheiro"
'retornar valores para linha 1
If Text3.Text <> 0 And Form407.Label17.Caption = "ENTREGA" And Check1.Value = 0 Then
Label6.Caption = "# LEVAR TROCO  " & Format(Text3.Text, "Currency")
Else
If Form407.Label17.Caption = "ENTREGA" Then
Label6.Caption = "# NAO LEVAR TROCO "
Else
        If Text3.Text <> 0 Then
        Label6.Caption = "# TROCO  " & Format(Text3.Text, "Currency")
        Else
        Label6.Caption = ""
        End If
End If
End If
Case "Cartão"
If Text2.Text <> 0 And Form407.Label17.Caption = "ENTREGA" And Check1.Value = 0 Then
Label6.Caption = "# LEVAR MAQUINA DE CARTAO PASSAR O VALOR DE  :  " & Format(Text2.Text, "Currency")
Else
Label6.Caption = "# NAO LEVAR MAQUINA !"
End If


Case "Ticket"
If Text2.Text <> 0 And Form407.Label17.Caption = "ENTREGA" And Check1.Value = 0 Then
Label6.Caption = "# RECEBER EM TICKET VALOR :  " & Format(Text2.Text, "Currency")
Else
Label6.Caption = "# VALO PAGO EM TICKET"
End If


Case "Anotar" '
'retornar valores para linha 4
If Text2.Text <> 0 And Form407.Label17.Caption = "ENTREGA" And Check1.Value = 0 Then
Label6.Caption = "# ANOTADO O VALOR  :  " & Format(Text2.Text, "Currency")
Else
Label6.Caption = "# ANOTADO !"

End If
 
 
End Select
SalvarPagamento

End Sub

Private Sub Text2_Click()
Text2.Text = ""
End Sub

Private Sub Text2_GotFocus()
If Text1.Visible = True Then
Text2.Text = ""
End If
End Sub

Private Sub Text2_LostFocus()
Dim valor1 As Single
Dim valor2 As Single
If Text2.Text <> "" Then
valor1 = Text2
Else
valor1 = 0
End If
If Form407.Label4.Caption = "label4" Then
Form407.Label4.Caption = Form407.Text1.Text
End If

valor2 = Form407.Label4.Caption
If valor1 > valor2 Then
MsgBox "Valor não pode ser maior que " & Form407.Label4.Caption, vbCritical, "ERRO"
Text2.Text = ""
Text2.SetFocus

End If
End Sub

Private Sub Text3_Change()
Label4.Caption = Replace(Text3.Text, ",", ".")
End Sub

Public Sub SalvarPagamento()
If Text4.Text <> 1 Then
ConServer
Dim numeroPeditoFiltrado As Integer
Dim EntregarBUscar As String
Dim formaPagamento As String
Dim numPedido As Integer




Dim ValorTotal As String
Dim ValorRecebido As String
Dim ValorPago As String
Dim Troco As String
Dim NumClient  As Integer
Dim Observacao As String
Dim ReceberPago As String

Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset
Set rs.ActiveConnection = con
EntregarBUscar = Form407.Label17.Caption
numPedido = Form407.Label1.Caption
formaPagamento = Format(Label2.Caption, ">")
ValorTotal = DBLcurre(Label7.Caption)
ReceberPago = Format(Label5.Caption, ">")
ValorRecebido = DBLcurre(Label1.Caption)
ValorPago = DBLcurre(Label3.Caption)
Troco = DBLcurre(Label4.Caption)
If Form2.Text11.Text = "inicio" Then
NumClient = 11
Else
NumClient = Form2.Text11.Text
End If
Observacao = Label6.Caption

rs.CursorLocation = adUseClient

 'apanhar numero do  PEDIDO OFICIAL
        
         sql = "INSERT INTO `robofi61_order_taker`.`Pagamento` (`EntregarBuscar`, `ReceberPago`, `FormaPagamento`, `ValorTotal`, `ValorRecebido`, `ValorPago`, `Troco`, `NUmPedido`, `NumClient`, `Observação`) VALUES (' " & EntregarBUscar & " ', ' " & ReceberPago & " ',' " & formaPagamento & " ',' " & ValorTotal & " ',' " & ValorRecebido & " ',' " & ValorPago & " ',' " & Troco & " ',' " & numPedido & " ',' " & NumClient & " ',' " & Observacao & " ' )"
         rs.Open sql

 Set rs = Nothing
End If
End Sub

