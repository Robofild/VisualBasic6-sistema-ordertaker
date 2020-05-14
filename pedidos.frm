VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form Form5 
   Caption         =   "Cardápio"
   ClientHeight    =   9855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12435
   LinkTopic       =   "Form5"
   ScaleHeight     =   9855
   ScaleWidth      =   12435
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Consultar "
      Height          =   975
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   11535
      Begin VB.CommandButton Command1 
         Caption         =   "Consultar"
         Height          =   375
         Left            =   10320
         TabIndex        =   7
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   5160
         TabIndex        =   6
         Top             =   360
         Width           =   4935
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Sabores"
         Height          =   255
         Left            =   3720
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Marca"
         Height          =   255
         Left            =   2520
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Nome"
         Height          =   255
         Left            =   1440
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Codigo"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   8895
      Left            =   360
      TabIndex        =   0
      Top             =   1440
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   15690
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
      Caption         =   "Massas "
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
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub DataGrid1_Click()
Form4.List1.AddItem ("ITEN DO CARDAPIO INCLUIDO")

End Sub
