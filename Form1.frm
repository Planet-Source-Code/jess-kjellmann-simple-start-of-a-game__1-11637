VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "3"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4920
   LinkTopic       =   "Form1"
   MousePointer    =   99  'Custom
   ScaleHeight     =   5865
   ScaleWidth      =   4920
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer6 
      Interval        =   10
      Left            =   3120
      Top             =   120
   End
   Begin VB.Timer Timer5 
      Interval        =   10
      Left            =   2520
      Top             =   120
   End
   Begin VB.Timer Timer4 
      Interval        =   10
      Left            =   1920
      Top             =   120
   End
   Begin VB.Timer Timer3 
      Interval        =   10
      Left            =   1320
      Top             =   120
   End
   Begin VB.Timer Timer2 
      Interval        =   10
      Left            =   720
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   120
      Top             =   120
   End
   Begin VB.Label Stick 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   960
      TabIndex        =   0
      Top             =   5700
      Width           =   1215
   End
   Begin VB.Shape Ball1 
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   1440
      Shape           =   3  'Circle
      Top             =   3000
      Width           =   105
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Thank you for downloading this simple gamecode.
' This only a start for a game, but it gives you the idea of how it works..
' You can use the code in your own game, but I would like if you would
' put my name someplace in your game... Or at least vote for me :o)
' Greetings from McGoat
' Sorry about the bad english, I'm from Denmark... :o)
' I program better that this, this is only for helping new ones
' So I have made this easy to understand

Public Way As String ' String that contains the way of the ball

Private Sub Form_Load()
newgame ' New Game
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Stick.Left = X - Stick.Width / 2 ' Make the Stick move after the mouse
End Sub

Private Sub Timer1_Timer()
Select Case Way: ' What way are the ball moving?
Case "southeast": ' SouthEast
Ball1.Top = Ball1.Top + 100 ' Make the ball go down
Ball1.Left = Ball1.Left + 100 ' Make the ball go right
Case "southwest": ' SouthWest
Ball1.Top = Ball1.Top + 100 ' Make the ball go down
Ball1.Left = Ball1.Left - 100 ' Make the ball go left
Case "northeast": ' NorthEast
Ball1.Top = Ball1.Top - 100 ' Make the ball go up
Ball1.Left = Ball1.Left + 100 ' Make the ball go right
Case "northwest": ' NorthWest
Ball1.Top = Ball1.Top - 100 ' Make the ball go up
Ball1.Left = Ball1.Left - 100 ' Make the ball go left
End Select ' End finding out.
End Sub

Private Sub Timer2_Timer()
' Check with Stick
If Ball1.Left > Stick.Left And Ball1.Left < Stick.Left + Stick.Width Then ' Is ball within Stick?
    If Ball1.Top = 5600 Then ' Is ball at hitpoint?
        If Ball1.Left < Stick.Left + (Stick.Width / 2) Then ' Is ball at left side at Stick?
            Way = "northwest" ' Go NorthWest
        Else ' Right Side
            Way = "northeast" ' Go NorthEast
        End If
    End If
End If
End Sub

Private Sub Timer3_Timer()
' Check with Rightwall
If Way = "northeast" Then ' Is way northeast?
If Ball1.Left + Ball1.Width >= Me.Width Then Way = "northwest" ' if ball hits wall then change direction to northwest
End If

If Way = "southeast" Then ' Is way southeast?
If Ball1.Left + Ball1.Width >= Me.Width Then Way = "southwest" ' if ball hits wall then change direction to southwest
End If
End Sub

Private Sub Timer4_Timer()
' Check with TopWall
If Way = "northwest" Then ' Is way northwest?
If Ball1.Top <= 0 Then Way = "southwest" ' if ball hits wall then change direction to southwest
End If

If Way = "northeast" Then ' Is way northeast?
If Ball1.Top <= 0 Then Way = "southeast" ' if ball hits wall then change direction to southeast
End If
End Sub

Private Sub Timer5_Timer()
' Check with Leftwall
If Way = "northwest" Then ' Is way northwest?
If Ball1.Left <= 0 Then Way = "northeast" ' if ball hits wall then change direction to northeast
End If

If Way = "southwest" Then ' Is way southwest?
If Ball1.Left <= 0 Then Way = "southeast" ' if ball hits wall then change direction to southeast
End If
End Sub

Private Sub Timer6_Timer()
' Life lost
If Ball1.Top >= Me.Height Then ' is ball under Stick?
Me.Caption = Me.Caption - 1 ' Life - 1
If Me.Caption = 0 Then ' If nomore lifes
DEAD1 ' DIE
End If
Ball1.Top = 3000 ' If still more lifes, then put ball at start
Ball1.Left = 1440
Way = "southeast"
End If
End Sub

Public Function DEAD1()
Timer1.Enabled = False ' Stop timer
Timer2.Enabled = False ' Stop timer
Timer3.Enabled = False ' Stop timer
Timer4.Enabled = False ' Stop timer
Timer5.Enabled = False ' Stop timer
Timer6.Enabled = False ' Stop timer
Me.Caption = "DEAD" ' Change Form Title To DEAD
Answer = MsgBox("New Game?", vbYesNo, "New Game?") ' Ask for new game
If Answer = "6" Then newgame ' If user answers YES then Make a new game
If Answer = "7" Then EXI ' If user answers NO then Exit Game
End Function

Public Function newgame()
Me.Caption = "3" ' 3 Lifes
Ball1.Left = 1440 ' Ball Start
Ball1.Top = 3000 ' Ball Start
Way = "southeast" ' Direction = Southeast
Timer1.Enabled = True ' Start Timer
Timer2.Enabled = True ' Start Timer
Timer3.Enabled = True ' Start Timer
Timer4.Enabled = True ' Start Timer
Timer5.Enabled = True ' Start Timer
Timer6.Enabled = True ' Start Timer
End Function

Public Function EXI()
End ' Exit Game
End Function
