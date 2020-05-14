VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.ocx"
Begin VB.Form Form502 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cancelamento"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10200
   Icon            =   "cancelamento.frx":0000
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   10200
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2775
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   8055
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   495
         Left            =   840
         TabIndex        =   1
         Top             =   2040
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   873
         _Version        =   327682
         BorderStyle     =   1
         Appearance      =   1
         Min             =   1e-4
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Cancelando o pedido:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   7695
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "lablel3"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   600
         Left            =   240
         TabIndex        =   2
         Top             =   1080
         Width           =   7695
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   600
      Top             =   3240
   End
End
Attribute VB_Name = "Form502"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub Timer1_Timer()
'Form501.Show
Dim indexB As Integer

  For indexB = 1 To 100
  ProgressBar1.Value = indexB / 2
        'Sleep (12)
                                        
  Next indexB

  ProgressBar1.Value = 100
'  cmdImprimir

  Timer1.Interval = 0
  MsgBox "Pedido cancelado com sucesso", , "Cancelado"
  Unload Me

End Sub
