VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form Form22 
   Caption         =   "Clientes cadastrados / Coincidência de telefone"
   ClientHeight    =   6150
   ClientLeft      =   4095
   ClientTop       =   4635
   ClientWidth     =   12780
   Icon            =   "clienteTelefones.frx":0000
   LinkTopic       =   "Form9"
   Moveable        =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   12780
   StartUpPosition =   1  'CenterOwner
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4935
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   8705
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   19
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
Attribute VB_Name = "Form22"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub DataGrid1_DblClick()
Form2.Visible = True

Form2.Text11.Text = DataGrid1.Columns(0).Value
Form2.Text12.Text = 0
Form22.Visible = False
'Form2.Command2.SetFocus
'Form2.DataCombo1.SetFocus
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Form2.Visible = True

Form2.Text11.Text = DataGrid1.Columns(0).Value
Form2.Text12.Text = 0
Form22.Visible = False
'Form2.Command2.SetFocus
'Form2.DataCombo1.SetFocus
End If
End Sub

Private Sub Form_Load()
Form2.Show
Form2.Visible = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'If Form2.Text4.Enabled = True And Form2.Text4.Visible = True Then
'Form2.DataCombo1.SetFocus
'End If
'Unload Form2
'Unload Form22
'Form3.Show
 'SendKeys "%{TAB}"
End Sub


