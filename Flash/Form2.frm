VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Enter the timer's intervall"
   ClientHeight    =   960
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   960
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   360
      Left            =   3270
      TabIndex        =   1
      Top             =   285
      Width           =   840
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   1005
      TabIndex        =   0
      Top             =   285
      Width           =   1755
   End
   Begin VB.Label Label1 
      Caption         =   "1000=1 sec"
      Height          =   165
      Left            =   1620
      TabIndex        =   2
      Top             =   720
      Width           =   1965
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo m:
Form1.Timer1.Interval = Val(Form2.Text1.Text)
Form1.HScroll1.Value = Val(Text1.Text)
Unload Me
Exit Sub
m:
MsgBox "Invalid input"
End Sub
