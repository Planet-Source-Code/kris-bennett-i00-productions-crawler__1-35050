VERSION 5.00
Begin VB.Form submit_score 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "High Scores"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7065
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   7065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   3240
      TabIndex        =   1
      Top             =   1440
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   450
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   3495
   End
End
Attribute VB_Name = "submit_score"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim response As Integer


Private Sub Command1_Click()
End
End Sub

Private Sub Form_Load()
Dim collision_loss, game_time, score
  collision_loss = Form2.collision_count.Text * Form2.collision_count.Text * Form2.collision_count.Text * Form2.collision_count.Text
  game_time = Form2.elapsed_time.Caption
  score = Val(collision_loss) + Val(game_time)
  response = MsgBox("you scored: " & score, vbInformation, "SCORE")

Dim tempstring1, tempstring2
  Open (App.Path & "\High Scores.txt") For Input As 1
  Do Until EOF(1)
    Line Input #1, tempstring1
    Line Input #1, tempstring2
    List1.AddItem tempstring1 & " - " & tempstring2
  Loop
  Close #1
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub
