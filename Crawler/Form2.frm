VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FF8080&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Crawler"
   ClientHeight    =   6030
   ClientLeft      =   1050
   ClientTop       =   1290
   ClientWidth     =   8175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   8175
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   7560
      ScaleHeight     =   345
      ScaleWidth      =   510
      TabIndex        =   6
      Top             =   5760
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.TextBox remaining_items 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3120
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Interval        =   4
      Left            =   480
      Top             =   -120
   End
   Begin VB.TextBox collision_count 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Text            =   "0"
      Top             =   120
      Width           =   735
   End
   Begin VB.Label cmd_leveldown 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "<"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   6240
      TabIndex        =   8
      Top             =   120
      Width           =   210
   End
   Begin VB.Label cmd_levelup 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   ">"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   7200
      TabIndex        =   7
      Top             =   120
      Width           =   210
   End
   Begin VB.Image Pac2 
      Height          =   255
      Left            =   0
      Top             =   480
      Width           =   255
   End
   Begin VB.Image imgObj 
      Height          =   180
      Index           =   10
      Left            =   2400
      Picture         =   "Form2.frx":0000
      Top             =   5760
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgObj 
      Height          =   240
      Index           =   9
      Left            =   2160
      Picture         =   "Form2.frx":0282
      Top             =   5760
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgObj 
      Height          =   255
      Index           =   8
      Left            =   1920
      Picture         =   "Form2.frx":05C4
      Top             =   5760
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image imgObj 
      Height          =   240
      Index           =   7
      Left            =   1680
      Picture         =   "Form2.frx":0A02
      Top             =   5760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgObj 
      Height          =   240
      Index           =   6
      Left            =   1440
      Picture         =   "Form2.frx":0F84
      Top             =   5760
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgObj 
      Height          =   240
      Index           =   5
      Left            =   1200
      Picture         =   "Form2.frx":12C6
      Top             =   5760
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgObj 
      Height          =   270
      Index           =   4
      Left            =   960
      Picture         =   "Form2.frx":1608
      Top             =   5760
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image imgObj 
      Height          =   240
      Index           =   3
      Left            =   720
      Picture         =   "Form2.frx":1A3A
      Top             =   5760
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgObj 
      Height          =   225
      Index           =   2
      Left            =   480
      Picture         =   "Form2.frx":1D7C
      Top             =   5760
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgObj 
      Height          =   225
      Index           =   1
      Left            =   240
      Picture         =   "Form2.frx":208E
      Top             =   5760
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgObj 
      Height          =   195
      Index           =   0
      Left            =   0
      Picture         =   "Form2.frx":23A0
      Top             =   5760
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Items Left"
      Height          =   255
      Left            =   2280
      TabIndex        =   5
      Top             =   120
      Width           =   855
   End
   Begin VB.Image item 
      Height          =   240
      Index           =   0
      Left            =   240
      Stretch         =   -1  'True
      Top             =   240
      Width           =   255
   End
   Begin VB.Label cmd_level 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   6465
      TabIndex        =   3
      Top             =   120
      Width           =   720
   End
   Begin VB.Label cmd_exit 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  EXIT  "
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   7440
      TabIndex        =   2
      Top             =   120
      Width           =   570
   End
   Begin VB.Label elapsed_time 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.Shape wall 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF8080&
      Height          =   255
      Index           =   0
      Left            =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WallAmount As Long
Dim ItemAmount As Long
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long

Function PlayWAV(filename As String, Optional SyncExec As Boolean) As Boolean
    Const SND_ASYNC = &H1
    If SyncExec Then
        ' play the file synchronously
        PlayWAV = PlaySound(filename, 0, 0)
    Else
        ' play the file asynchronously
        PlayWAV = PlaySound(filename, 0, SND_ASYNC)
    End If
End Function

Private Function FileExists(FullFileName As String) As Boolean
    On Error GoTo MakeF
        'If file does Not exist, there will be an Error
        Open FullFileName For Input As #1
        Close #1
        'no error, file exists
        FileExists = True
    Exit Function
MakeF:
        'error, file does Not exist
        FileExists = False
    Exit Function
End Function

Function GoToMap(MapNo As Integer)
cmd_level.Caption = MapNo
For i = 1 To item.Count - 1
    Unload item(i)
Next
For i = 1 To wall.Count - 1
    Unload wall(i)
Next
Form_Load
End Function


Private Sub cmd_levelup_Click()

GoToMap cmd_level.Caption + 1

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

Dim move As Boolean, i As Integer
move = True
'The only outcome from this "case" code is
Select Case KeyCode
  Case vbKeyDown
  Pac2.Picture = LoadPicture(App.Path & "\down.bmp")
      For i = 1 To WallAmount
      If wall(i).Top = (Pac2.Top + 240) And wall(i).Left = Pac2.Left Then
        'what happens when hitting the wall
        collision_count.Text = Val(collision_count.Text) + 1
        wall(i).BackColor = &HC0C0FF
        move = False
      End If
      Next
    If move = True Then
        Pac2.Top = Pac2.Top + 240
      End If
  
  Case vbKeyUp
  Pac2.Picture = LoadPicture(App.Path & "\up.bmp")
      For i = 0 To WallAmount
      If wall(i).Top = (Pac2.Top - 240) And wall(i).Left = Pac2.Left Then
        'what happens when hitting the wall
        collision_count.Text = Val(collision_count.Text) + 1
        wall(i).BackColor = &HC0C0FF
            move = False
      End If
      Next
      If move = True Then
        Pac2.Top = Pac2.Top - 240
      End If
  
  Case vbKeyRight
  Pac2.Picture = LoadPicture(App.Path & "\right.bmp")
      For i = 0 To WallAmount
      If wall(i).Top = Pac2.Top And wall(i).Left = (Pac2.Left + 240) Then
        'what happens when hitting the wall
        collision_count.Text = Val(collision_count.Text) + 1
        wall(i).BackColor = &HC0C0FF
        move = False
      End If
      Next
      If move = True Then
        Pac2.Left = Pac2.Left + 240
      End If

  Case vbKeyLeft
  Pac2.Picture = LoadPicture(App.Path & "\left.bmp")
      For i = 0 To WallAmount
      If wall(i).Top = Pac2.Top And wall(i).Left = (Pac2.Left - 240) Then
        'what happens when hitting the wall
        collision_count.Text = Val(collision_count.Text) + 1
        wall(i).BackColor = &HC0C0FF
        move = False
      End If
      Next
      If move = True Then
        Pac2.Left = Pac2.Left - 240
      End If

End Select

item_collision

If Pac2.Left < 0 Then
    Pac2.Left = 7920
End If
If Pac2.Left > 7920 Then
    Pac2.Left = 0
End If

If Pac2.Top < 480 Then
    Pac2.Top = 5760
End If
If Pac2.Top > 5760 Then
    Pac2.Top = 480
End If

End Sub

Function item_collision()
On Error Resume Next
Dim i
      For i = 1 To ItemAmount
      If item(i).Top = Pac2.Top And item(i).Left = Pac2.Left And item(i).Visible = True Then
        'what happens when on an item
        item(i).Visible = False
        PlayWAV App.Path & "\newalert.wav", False
      Else
        'when not on an object
      End If
      Next
      
'checking mow many icons / items are left
Dim item_visible_count
item_visible_count = 0
      For i = 1 To ItemAmount
      If item(i).Visible = True Then
        item_visible_count = Val(item_visible_count) + 1
      End If
      remaining_items.Text = item_visible_count
      Next
      
      If remaining_items.Text = "0" Then
        'when all items are collected
        'Timer1.Enabled = False
        MsgBox "You compleated level " & cmd_level.Caption & " and scored: " & elapsed_time, vbInformation, "SCORE"
        If cmd_levelup.Enabled = True Then
            cmd_levelup_Click
            elapsed_time.Caption = 0
        Else
            MsgBox "Congrats.. u compleated Crawler"
            End
        End If
      End If
      
      
End Function

Private Sub Form_Load()

If FileExists(App.Path & "\map" & cmd_level.Caption - 1 & ".bmp") = False Then
    cmd_leveldown.Enabled = False
Else
    cmd_leveldown.Enabled = True
End If

If FileExists(App.Path & "\map" & cmd_level.Caption + 1 & ".bmp") = False Then
    cmd_levelup.Enabled = False
Else
    cmd_levelup.Enabled = True
End If

Randomize
Picture1.Picture = LoadPicture(App.Path & "\map" & cmd_level.Caption & ".bmp")

For w = 1 To Picture1.Width Step 15
    For h = 1 To Picture1.Height Step 15
        If Picture1.Point(w, h) = vbBlue Then
            Pac2.Left = w * 16 - 16
            Pac2.Top = h * 16 + 480 - 16
            Pac2.Picture = LoadPicture(App.Path & "\up.bmp")
        End If
        
        If Picture1.Point(w, h) = vbRed Then
            Load item(item.Count)
            item(item.Count - 1).Left = w * 16 - 16
            item(item.Count - 1).Top = h * 16 + 480 - 16
            item(item.Count - 1).Visible = True
            item(item.Count - 1).Picture = imgObj(Int(Rnd * 10)).Picture
        End If
        If Picture1.Point(w, h) = vbBlack Then
            Load wall(wall.Count)
            wall(wall.Count - 1).Left = w * 16 - 16
            wall(wall.Count - 1).Top = h * 16 + 480 - 16
            wall(wall.Count - 1).Visible = True
        End If
    Next
Next
WallAmount = wall.Count - 1
ItemAmount = item.Count - 1

Me.Visible = True
If cmd_level.Caption = 1 Then MsgBox "Once the game starts so will the timer, get ready for the action", vbOKOnly, ""

remaining_items.Text = ItemAmount



End Sub

Private Sub Form_Unload(Cancel As Integer)
If MsgBox("Really Exit?", vbOKCancel) = vbOK Then
    End
Else
    Cancel = 1
End If
End Sub

Private Sub cmd_exit_Click()
Unload Me
End Sub

Private Sub cmd_leveldown_Click()

GoToMap cmd_level.Caption - 1

End Sub

Private Sub Timer1_Timer()
elapsed_time.Caption = Val(elapsed_time.Caption) + 1
End Sub
