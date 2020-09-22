VERSION 5.00
Begin VB.Form frmBall 
   Caption         =   "Gravity"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrDropBall 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   2160
   End
   Begin VB.Timer tmrPickUpBall 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   2640
   End
   Begin VB.Label lblGravity 
      Caption         =   "0.15"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblBlue 
      Caption         =   "0"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblGreen 
      Caption         =   "255"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblRed 
      Caption         =   "0"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape ground 
      BackColor       =   &H80000003&
      BackStyle       =   1  'Opaque
      Height          =   1815
      Left            =   0
      Top             =   6720
      Width           =   11895
   End
   Begin VB.Shape ball 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   960
      Shape           =   3  'Circle
      Top             =   6240
      Width           =   495
   End
   Begin VB.Menu file 
      Caption         =   "File"
      Begin VB.Menu properties 
         Caption         =   "Properties..."
      End
      Begin VB.Menu restart 
         Caption         =   "Restart"
      End
      Begin VB.Menu quit 
         Caption         =   "Quit"
      End
   End
   Begin VB.Menu help 
      Caption         =   "Help"
      Begin VB.Menu howtoplay 
         Caption         =   "How to Play"
      End
      Begin VB.Menu about 
         Caption         =   "About..."
      End
   End
End
Attribute VB_Name = "frmBall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MouseX As Integer
Dim MouseY As Integer
Dim hJump As Integer
Dim Ground_Level As Integer
Dim Max_Height As Integer
Dim BallGoesUp As Boolean
Dim gravity As Double
Dim BallRed As Integer
Dim BallGreen As Integer
Dim BallBlue As Integer
Private Sub Form_Load()
    BallRed = 0
    BallGreen = 255
    BallBlue = 0
    hJump = 0
    ball.BackColor = RGB(BallRed, BallGreen, BallBlue)
    Ground_Level = ground.Top
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    gravity = lblGravity.Caption
    If Button = 1 Then
        If X >= ball.Left And X <= (ball.Left + ball.Width) And Y >= ball.Top And Y <= (ball.Top + ball.Height) Then
            tmrDropBall.Enabled = False
            tmrPickUpBall.Enabled = True
        End If
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseX = X
    MouseY = Y
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If X >= ball.Left And X <= (ball.Left + ball.Width) And Y >= ball.Top And Y <= (ball.Top + ball.Height) Then
            tmrPickUpBall.Enabled = False
            Max_Height = Ground_Level - ball.Top + 100
            BallGoesUp = False
            tmrDropBall.Enabled = True
        End If
    End If
End Sub

Private Sub Label3_Click()
    ball.Left = 1440
    ball.Top = 6240
End Sub

Private Sub howtoplay_Click()
    MsgBox ("Drag and drop the ball to move it.  Watch it fall.  Repeat.")
End Sub

Private Sub Properties_Click()
    Prop.txtGravity.Text = gravity * 30
    lblGravity.Caption = gravity
    Prop.BallColor.BackColor = ball.BackColor
    Prop.Show
End Sub

Private Sub quit_Click()
    End
End Sub

Private Sub tmrCollision_Timer()
    
    Dim Ratio1 As Integer
    Dim Ratio2 As Integer
    
    Ratio1 = (ball.Left - line1.X1) / (ball.Top + ball.Height - line1.Y1)
    Ratio2 = (line1.X2 - line1.X1) / (line1.Y2 - line1.Y1)
    
    If Ratio1 = Ratio2 And ball.Left >= line1.X1 And ball.Left + ball.Width <= line1.X2 Then
        Ground_Level = ball.Top + ball.Height
        hJump = 100
        tmrDropBall.Enabled = True
    End If
End Sub

Private Sub tmrDropBall_Timer()
If BallGoesUp = True Then
    If (Max_Height <= ball.Height + 5) Then
        tmrDropBall.Enabled = False
    ElseIf (ball.Top - gravity * (Max_Height - (ball.Top - (Ground_Level - Max_Height)))) < Ground_Level - Max_Height Then
        BallGoesUp = False
    Else
        ball.Top = ball.Top - gravity * (ball.Top - (Ground_Level - Max_Height))
        ball.Left = ball.Left + hJump
    End If
Else        'ball goes down
    If (ball.Top + gravity * (Max_Height - (ball.Top - (Ground_Level - Max_Height)))) > Ground_Level - ball.Height Then
        ball.Top = Ground_Level - ball.Height
        BallGoesUp = True
        Max_Height = (Max_Height / 1.5)
    Else
        ball.Top = ball.Top + gravity * (ball.Top - (Ground_Level - Max_Height))
        ball.Left = ball.Left + hJump
    End If
End If
End Sub

Private Sub tmrPickUpBall_Timer()
    ball.Top = MouseY - (ball.Height / 2)
    ball.Left = MouseX - (ball.Width / 2)
End Sub
