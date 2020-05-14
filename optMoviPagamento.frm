VERSION 5.00
Begin VB.Form Form406_10 
   Caption         =   "CAIXA"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "optMoviPagamento.frx":0000
   LinkTopic       =   "Form9"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF8080&
      Caption         =   "F1-    a pagar"
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
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   480
      Width           =   3735
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H0080FF80&
      Caption         =   "F2-        Pago!"
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
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1680
      Width           =   3735
   End
End
Attribute VB_Name = "Form406_10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
'retirar o frete
'BALCAO
Form407.Show
Form407.Label18.Caption = "r"
End Sub


Private Sub Command5_Click()

'retirar o frete
'BALCAO
Form407.Show
Form407.Label18.Caption = "Pago"


End Sub
