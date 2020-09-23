VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Window Flasher"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4065
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4065
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Custom"
      Height          =   330
      Left            =   3120
      TabIndex        =   7
      Top             =   780
      Width           =   885
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Option1"
      Height          =   225
      Left            =   3390
      TabIndex        =   4
      Top             =   1710
      Width           =   255
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   225
      Left            =   3360
      TabIndex        =   3
      Top             =   2295
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   480
      TabIndex        =   2
      Top             =   360
      Width           =   2340
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      Left            =   105
      Max             =   3000
      Min             =   500
      TabIndex        =   0
      Top             =   765
      Value           =   500
      Width           =   2940
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3195
      Top             =   270
   End
   Begin VB.Label Label3 
      Caption         =   "When window is minimized the shape on the taskbar change color."
      Height          =   465
      Left            =   75
      TabIndex        =   6
      Top             =   2250
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "Flashes the window activate /deactivate"
      Height          =   465
      Left            =   135
      TabIndex        =   5
      Top             =   1680
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Gibreel"
      Height          =   285
      Left            =   1095
      TabIndex        =   1
      Top             =   1425
      Width           =   2610
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public boulis As Boolean
Private Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long

Private Sub Command1_Click()
Form2.Show
End Sub

Private Sub Form_Load()
Text1.Text = 0
End Sub

Private Sub HScroll1_Change()
Text1.Text = "Intervall=" & HScroll1.Value
End Sub

Private Sub HScroll1_Scroll()
Text1.Text = "Intervall=" & HScroll1.Value

End Sub

Private Sub Option1_Click()
boulis = False
End Sub

Private Sub Option2_Click()
boulis = True
End Sub

Private Sub Timer1_Timer()
FlashWindow Me.hwnd, boulis
End Sub

