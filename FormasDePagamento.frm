VERSION 5.00
Begin VB.Form Form408 
   BackColor       =   &H00FF0000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Decidindo forma de pagamento"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3930
   Icon            =   "FormasDePagamento.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form14"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   3930
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "F4-       Anotar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   3360
      Width           =   3735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "F3-      Ticket"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   3735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "F2-     Cartão"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "F1-  Dinheiro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "Form408"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()


Form409.Show
'Form409.Text2 = Form407.Text1
Form409.Label2 = "Dinheiro"
AcenderFrames
End Sub

Private Sub Command2_Click()

Form409.Show
'Form409.Text2 = Form407.Text1
Form409.Label2 = "Cartão"
ApagarFrames
End Sub

Private Sub Command3_Click()

Form409.Show
'Form409.Text2 = Form407.Text1
Form409.Label2 = "Ticket"
ApagarFrames
End Sub

Private Sub Command4_Click()

Form409.Show
'Form409.Text2 = Form407.Text1
Form409.Label2 = "Anotar"
Form409.Check1.Value = 1
ApagarFrames
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF1 Then
  Call Command1_Click
  End If
    If KeyCode = vbKeyF2 Then
   Call Command2_Click
  End If
    If KeyCode = vbKeyF3 Then
    Call Command3_Click
  End If
    If KeyCode = vbKeyF4 Then
    
    Call Command4_Click
    
  End If
End Sub




Public Sub ApagarFrames()
Form409.Text1.Visible = False
Form409.Text3.Visible = False
Form409.Label12.Visible = False
Form409.Label10.Visible = False
End Sub
Public Sub AcenderFrames()
Form409.Text1.Visible = True
Form409.Text3.Visible = True
Form409.Label12.Visible = True
Form409.Label10.Visible = True
End Sub
