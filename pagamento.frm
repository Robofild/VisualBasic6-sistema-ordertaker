VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "Pagamento"
   ClientHeight    =   4530
   ClientLeft      =   5115
   ClientTop       =   5250
   ClientWidth     =   12015
   Icon            =   "pagamento.frx":0000
   LinkTopic       =   "Form6"
   ScaleHeight     =   4530
   ScaleWidth      =   12015
   Begin VB.CommandButton Command2 
      Caption         =   "Finalizar "
      Height          =   735
      Left            =   7440
      TabIndex        =   5
      Top             =   3120
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Troco ?"
      Height          =   495
      Left            =   7080
      TabIndex        =   4
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   3
      Top             =   2040
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   2
      Text            =   "75,00"
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   9120
      Top             =   1560
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "pagamento.frx":000C
      Left            =   2520
      List            =   "pagamento.frx":0016
      TabIndex        =   0
      Text            =   "Forma de pagamento"
      Top             =   600
      Width           =   3255
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
