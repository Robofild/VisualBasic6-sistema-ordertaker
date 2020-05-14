VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Pedido"
   ClientHeight    =   9105
   ClientLeft      =   4065
   ClientTop       =   3255
   ClientWidth     =   13920
   Icon            =   "cPedidos.frx":0000
   LinkTopic       =   "Form4"
   ScaleHeight     =   9105
   ScaleWidth      =   13920
   Begin VB.ListBox List1 
      Height          =   3375
      ItemData        =   "cPedidos.frx":000C
      Left            =   720
      List            =   "cPedidos.frx":000E
      TabIndex        =   17
      Top             =   5040
      Width           =   8535
   End
   Begin VB.ComboBox Combo9 
      Height          =   315
      Left            =   10320
      TabIndex        =   16
      Text            =   "Combo2"
      Top             =   720
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Salvar"
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
      Left            =   9840
      TabIndex        =   15
      Top             =   6720
      Width           =   3735
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "cPedidos.frx":0010
      Left            =   4440
      List            =   "cPedidos.frx":0020
      TabIndex        =   14
      Text            =   "Produtos"
      Top             =   720
      Width           =   4455
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H00808000&
      Caption         =   "Retirar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   11640
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5160
      Width           =   1815
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H0080FF80&
      Caption         =   "Acressimo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5160
      Width           =   1815
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H008080FF&
      Enabled         =   0   'False
      Height          =   1335
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3120
      Width           =   1815
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00808080&
      Enabled         =   0   'False
      Height          =   1335
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3120
      Width           =   1815
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H0000FF00&
      Enabled         =   0   'False
      Height          =   1335
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3120
      Width           =   1815
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00004040&
      Caption         =   "Fatias"
      Enabled         =   0   'False
      Height          =   1335
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3120
      Width           =   1815
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FF00FF&
      Enabled         =   0   'False
      Height          =   1335
      Left            =   11640
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FF8080&
      Enabled         =   0   'False
      Height          =   1335
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Enabled         =   0   'False
      Height          =   1335
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080FFFF&
      Enabled         =   0   'False
      Height          =   1335
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "Texto"
      Enabled         =   0   'False
      Height          =   1335
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cardapio"
      Enabled         =   0   'False
      Height          =   1335
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1440
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Codigo do produto"
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
   Begin VB.Menu mnu_configuracao 
      Caption         =   "Configuração"
      Begin VB.Menu Mnu_personalize 
         Caption         =   "Personalização"
      End
      Begin VB.Menu MnuCadastro 
         Caption         =   "Cadasto"
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MyArray(6, 3)
Dim ingrediente1, ingrediente2, ingrediente3, ingrediente4, ingrediente5, ingrediente6 As Integer
Dim ingrediente7, ingrediente8, ingrediente9, ingrediente10, ingrediente11, ingrediente12 As String
Dim Gerallista As String
Dim contadordeinice As Integer




Public Sub sanduiche()
Command1.Enabled = True
Command3.Caption = "Saladas"
Command4.Caption = "Frios"
Command5.Caption = "Defumados"
Command6.Caption = "Laticínios"

Command8.Caption = "Carnes"
Command9.Caption = "Pão"
Command10.Caption = "Molho"





End Sub
Public Sub Massas()
Command1.Enabled = True

Command3.Caption = "Saladas"
Command4.Caption = "Frios"
Command5.Caption = "bordas"
Command6.Caption = "Laticínios"
Command7.Caption = "Fatias"
Command8.Caption = "Molho"
Command9.Caption = "Defumados"
Command10.Caption = ""

Command11.Enabled = False
Command12.Enabled = False



End Sub


Private Sub Combo1_Click()


Select Case Combo1.Text

Case "Sanduíches"

sanduiche
 ' Atendimento
Case "Massas"
Massas

'cadastro
Case "Bebidas"
Bebidas

 ' pedido
Case "Extras"
Extras

 
 
End Select
 Form1.Toolbar1.Buttons(3).Value = tbrPressed
'Massas
'Bebidas
'Extras

End Sub
Public Sub Bebidas()
Command1.Enabled = True
AbilitarTodos
Command3.Caption = "Refrigerante"
Command4.Caption = "Destiladas"
Command5.Caption = "Sucos"
Command6.Caption = "Cervejas"

Command8.Caption = "Sorvetes"
Command9.Caption = ""
Command10.Caption = ""
Command11.Enabled = False
Command12.Enabled = False
ingrediente12 = "Bebidas"





End Sub

Public Sub Extras()
AbilitarTodos
Command1.Enabled = True
'AbilitarTodos
Command3.Caption = "Porções"
Command4.Caption = "Saladas"
Command5.Caption = "Pratos"
Command6.Caption = "sopas"
Command7.Caption = "Batata Suiça"
Command8.Caption = ""
Command9.Caption = ""
Command10.Caption = ""
Command11.Enabled = False
Command12.Enabled = False
ingrediente12 = "Bebidas"





End Sub

Private Sub Command1_Click()
Command11.Enabled = True
Command12.Enabled = True
Form5.Show
Form5.DataGrid1.Caption = Combo1.Text


End Sub

Private Sub Command10_Click()
Form5.Show
Form5.DataGrid1.Caption = Command10.Caption
End Sub

Private Sub Command11_Click()
Gerallista = List1.List(0)

ingrediente12 = "Acressimo"
List1.Clear
List1.AddItem (Gerallista & " + " & ingrediente12)

AbilitarTodos
End Sub

Private Sub Command12_Click()

Gerallista = List1.List(0)

ingrediente12 = "Retirar"
List1.Clear
List1.AddItem (Gerallista & " + " & ingrediente12)

AbilitarTodos




End Sub

Private Sub Command13_Click()
'MsgBox (ingrediente12 & ":" & ingrediente1 & ":" & ingrediente12 & ":" & ingrediente2 & ":" & ingrediente12 & ":" & ingrediente3 & " : " & ingrediente12 & ":" & ingrediente4 & _
'ingrediente12 & ":" & ingrediente5 & ":" & ingrediente12 & ":" & ingrediente6 & ":" & ingrediente12 & ":" & ingrediente7 & ":" & ingrediente12 & ":" & ingrediente8 & ":" & ingrediente12 & ":" & ingrediente9 & _
 'ingrediente12 & ":" & ingrediente10)
 If (List1.List(0) <> "") Then
Form4.Hide
Form6.Show
Form1.Toolbar1.Buttons(3).Value = tbrPressed
Unload Me
Else
MsgBox "Você ainda não adicionou um pedido !", vbCritical, "Conclua o pedido!"


End If

End Sub

Private Sub Command2_Click()
ingrediente4 = InputBox("Entre com a quantidade de " & Combo1.Text, "Quantidade do produto")

Gerallista = List1.List(0)

ingrediente12 = "quantidade"
List1.Clear
List1.AddItem (Gerallista & " + " & ingrediente4)



End Sub

Public Sub AbilitarTodos()

Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
Command7.Enabled = True
Command8.Enabled = True
Command9.Enabled = True
Command10.Caption = ""





End Sub

Private Sub Command3_Click()
Form5.Show
Form5.DataGrid1.Caption = Command3.Caption
End Sub

Private Sub Command4_Click()
Form5.Show
Form5.DataGrid1.Caption = Command4.Caption
End Sub

Private Sub Command5_Click()
Form5.Show
Form5.DataGrid1.Caption = Command5.Caption
End Sub

Private Sub Command6_Click()
Form5.Show
Form5.DataGrid1.Caption = Command6.Caption
End Sub

Private Sub Command7_Click()
ingrediente5 = InputBox("Entre com a quantidade de Fatias  " & Combo1.Text, "Quantidade do produto")

Gerallista = List1.List(0)

ingrediente12 = "fatias"
List1.Clear
List1.AddItem (Gerallista & " + " & ingrediente5)

End Sub

Private Sub Command8_Click()
Form5.Show
Form5.DataGrid1.Caption = Command8.Caption
End Sub

Private Sub Command9_Click()
Form5.Show
Form5.DataGrid1.Caption = Command9.Caption
End Sub

Private Sub Mnu_personalize_Click()
Form444.Show
End Sub
