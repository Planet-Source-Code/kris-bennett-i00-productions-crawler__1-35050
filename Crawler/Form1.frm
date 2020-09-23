VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   3255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6555
   LinkTopic       =   "Form2"
   ScaleHeight     =   3255
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5280
      Top             =   1320
   End
   Begin VB.Image Image1 
      Height          =   3255
      Left            =   0
      Picture         =   "Form1.frx":0000
      Top             =   0
      Width           =   6555
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
CurrentMap = 1
End Sub

Private Sub Timer1_Timer()
Form2.Show
Unload Me
End Sub
