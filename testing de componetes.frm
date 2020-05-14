VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Form101 
   Caption         =   "Form9"
   ClientHeight    =   7350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14085
   LinkTopic       =   "Form9"
   ScaleHeight     =   7350
   ScaleWidth      =   14085
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text11 
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Text            =   "C:\Order_Taker\FINDFILE.AVI"
      Top             =   0
      Width           =   4215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   2160
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Animation1.Open = ""C:\Order_Taker\FINDFILE.AVI"""
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   1575
   End
   Begin MSComCtl2.Animation Animation1 
      Height          =   3735
      Left            =   3960
      TabIndex        =   0
      Top             =   1800
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   6588
      _Version        =   393216
      FullWidth       =   385
      FullHeight      =   249
   End
End
Attribute VB_Name = "Form101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Command1_Click()
'Set Command1.Picture = Picture1.Image

Animation1.Open Text11.Text

Animation1.AutoPlay = True
End Sub

Private Sub Command2_Click()
Animation1.Close
End Sub

