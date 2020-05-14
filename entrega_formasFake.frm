VERSION 5.00
Begin VB.Form Form406_1 
   Caption         =   "Forma de Entrega"
   ClientHeight    =   2940
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4350
   Icon            =   "entrega_formasFake.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form9"
   ScaleHeight     =   2940
   ScaleWidth      =   4350
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "F2-           BALCÃO"
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
      Left            =   240
      TabIndex        =   1
      Top             =   1560
      Width           =   3735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "F1-     ENTREGAR"
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
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "nao"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   855
   End
End
Attribute VB_Name = "Form406_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Form401.Text12.Text = "R$ 0,00"
Form2.Show
Form406_1.Hide
End Sub

Private Sub Command2_Click()
'verificar se e entrega ou buscar



End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF1 Then
   Call Command2_Click
  End If
    If KeyCode = vbKeyF2 Then
   Call Command1_Click
  End If
 
    
 
End Sub
