VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.ocx"
Begin VB.Form Form500 
   Caption         =   " Aguarde por favor!"
   ClientHeight    =   1710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8595
   Icon            =   "Aguarde.frx":0000
   LinkTopic       =   "Form9"
   ScaleHeight     =   1710
   ScaleWidth      =   8595
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   5040
      Top             =   360
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label C 
      Caption         =   "Carregando as informações "
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
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3855
   End
End
Attribute VB_Name = "Form500"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Sub Form_Load()
Timer1.Interval = 200
End Sub
Private Sub Timer1_Timer()
Dim indexB As Integer

'CommonDialog1.CancelError = True
'trarar erro
On Error GoTo error
For indexB = 0 To 100
  ProgressBar1.Value = indexB / 2
        'Sleep (6)
                                        
  Next indexB
 
  ProgressBar1.Value = 100
  Form500.Hide
  Timer1.Interval = 0

Exit Sub

error:
For indexB = 0 To 100
  ProgressBar1.Value = indexB / 2
        'Sleep (6)
                                        
  Next indexB
 
  ProgressBar1.Value = 100
  Form500.Hide
  Timer1.Interval = 0

Exit Sub

  
 
End Sub
