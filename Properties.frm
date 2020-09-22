VERSION 5.00
Begin VB.Form Prop 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Properties"
   ClientHeight    =   3195
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   4680
      TabIndex        =   14
      Top             =   600
      Width           =   1215
   End
   Begin VB.Timer tmrProperties 
      Interval        =   100
      Left            =   4560
      Top             =   2520
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   4320
      TabIndex        =   12
      Text            =   "255"
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   3240
      TabIndex        =   10
      Text            =   "255"
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2040
      TabIndex        =   8
      Text            =   "255"
      Top             =   1680
      Width           =   495
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Pre-Defined"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   1200
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Custom"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Properties.frx":0000
      Left            =   1560
      List            =   "Properties.frx":000D
      TabIndex        =   5
      Text            =   "Pick a color"
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox txtGravity 
      Height          =   405
      Left            =   2760
      TabIndex        =   2
      Text            =   "5"
      Top             =   240
      Width           =   375
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Shape BallColor 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   3120
      Shape           =   3  'Circle
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "Blue:"
      Height          =   255
      Left            =   3840
      TabIndex        =   13
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "Green:"
      Height          =   255
      Left            =   2640
      TabIndex        =   11
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Red:"
      Height          =   255
      Left            =   1560
      TabIndex        =   9
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Ball Color:"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label lblGravity 
      Caption         =   "Gravity:"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "Prop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim Red
Dim Green
Dim Blue

Private Sub CancelButton_Click()
    frmBall.ball.BackColor = RGB(0, 255, 0)
    Unload Me
End Sub

Private Sub cmdApply_Click()
    If Option1.Value = True Then
        frmBall.ball.BackColor = RGB(Text2.Text, Text3.Text, Text4.Text)
    Else
        If Combo1.Text = "Red" Then
            frmBall.ball.BackColor = RGB(255, 0, 0)
        ElseIf Combo1.Text = "Green" Then
            frmBall.ball.BackColor = RGB(0, 255, 0)
        ElseIf Combo1.Text = "Blue" Then
            frmBall.ball.BackColor = RGB(0, 0, 255)
        End If
    End If
End Sub

Private Sub Combo1_LostFocus()
    If Combo1.Text = "Red" Then
        BallColor.BackColor = RGB(255, 0, 0)
    ElseIf Combo1.Text = "Green" Then
        BallColor.BackColor = RGB(0, 255, 0)
    ElseIf Combo1.Text = "Blue" Then
        BallColor.BackColor = RGB(0, 0, 255)
    End If
End Sub


Private Sub Form_Load()
    Red = frmBall.lblRed.Caption
    Green = frmBall.lblGreen.Caption
    Blue = frmBall.lblBlue.Caption
    Text2.Text = Red
    Text3.Text = Green
    Text4.Text = Blue
End Sub

Private Sub OKButton_Click()
    Dim grav As Double
    If Option1.Value = True Then
        frmBall.ball.BackColor = RGB(Text2.Text, Text3.Text, Text4.Text)
    Else
        If Combo1.Text = "Red" Then
            Red = 255
            Blue = 0
            Green = 0
            frmBall.ball.BackColor = RGB(255, 0, 0)
        ElseIf Combo1.Text = "Green" Then
            Red = 0
            Blue = 0
            Green = 255
            frmBall.ball.BackColor = RGB(0, 255, 0)
        ElseIf Combo1.Text = "Blue" Then
            Red = 0
            Blue = 255
            Green = 0
            frmBall.ball.BackColor = RGB(0, 0, 255)
        End If
    End If
    grav = txtGravity.Text / 30
    frmBall.lblGravity.Caption = grav
    frmBall.lblBlue.Caption = Blue
    frmBall.lblGreen.Caption = Green
    frmBall.lblRed.Caption = Red
    Unload Me
End Sub

Private Sub Text2_Change()
    If Text2.Text < 0 Or Text2.Text > 255 Then
        MsgBox ("Error. Number must be between 0 and 255")
        Text2.Text = 0
    End If
End Sub

Private Sub Text3_Change()
    If Text3.Text < 0 Or Text3.Text > 255 Then
        MsgBox ("Error. Number must be between 0 and 255")
        Text3.Text = 0
    End If
End Sub

Private Sub Text4_Change()
    If Text4.Text < 0 Or Text4.Text > 255 Then
        MsgBox ("Error. Number must be between 0 and 255")
        Text4.Text = 0
    End If
End Sub

Private Sub tmrProperties_Timer()
If Option1.Value = True Then
    Red = Text2.Text
    Green = Text3.Text
    Blue = Text4.Text
    BallColor.BackColor = RGB(Red, Green, Blue)
    Combo1.Enabled = False
    Text2.Enabled = True
    Text3.Enabled = True
    Text4.Enabled = True
    Label3.Enabled = True
    Label4.Enabled = True
    Label5.Enabled = True
ElseIf Option2.Value = True Then
    Combo1.Enabled = True
    Text2.Enabled = False
    Text3.Enabled = False
    Text4.Enabled = False
    Label3.Enabled = False
    Label4.Enabled = False
    Label5.Enabled = False
End If
End Sub

Private Function CollisionDetect(src As Shape, tgt As Shape) As Boolean

'This function coded by Goose, (c) 2000
'This function checks to see if two objects, src and tgt, have
'collided. It works by analyzing the left, right, top,
'and bottom sides of the src object, and seeing if they have
'penetrated the borders of the tgt object.

'It goes something like this:
'Src Left(SL) = src.left
'Src Right(SR) = src.left + src.width
'Src Top(ST) = src.top
'Src Bottom(SB) = src.top + src.height
'Same goes for the target side (TL, TR, TT, TB)

'Then, side collision is detected as follows:
'Left Collision (L) = SL > TL and SL < TR
'Right Collision (R) = SR > TL and SR < TR
'Top Collision (T) = ST > TT and ST < TB
'Bottom Collision (B) = SB > TT and SB < TB

'Collision constitutes both a left/right penetration and a
'top/bottom penetration, or in other words:
'   LT + RT + LB + RB
'  T(L + R) + B(L + R)
'     (T + B)(L + R)
'The giant if statement is the expansion of this boolean function.

'Left(L): (src.left > tgt.left) and (src.left < (tgt.left + tgt.width))
'Right(R): ((src.left + src.width) > tgt.left) and ((src.left + src.width) < (tgt.left + tgt.width))
'Top(T): (src.top > tgt.top) and (src.top < (tgt.top + tgt.height))
'Bottom(B): ((src.top + src.height) > tgt.top) and ((src.top + src.height) < (tgt.top + tgt.height))
    If (((src.Left > tgt.Left) And (src.Left < (tgt.Left + tgt.Width))) Or (((src.Left + src.Width) > tgt.Left) And ((src.Left + src.Width) < (tgt.Left + tgt.Width)))) And (((src.Top > tgt.Top) And (src.Top < (tgt.Top + tgt.Height))) Or (((src.Top + src.Height) > tgt.Top) And ((src.Top + src.Height) < (tgt.Top + tgt.Height)))) Then
        CollisionDetect = True
    Else
        CollisionDetect = False
    End If
End Function

