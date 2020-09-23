VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "The Stupidest Game in the World"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer3 
      Interval        =   1
      Left            =   4200
      Top             =   1080
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   4200
      Top             =   600
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4200
      Top             =   120
   End
   Begin VB.Label lblEnd 
      AutoSize        =   -1  'True
      Caption         =   "GAME OVER"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   840
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   2955
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      Caption         =   "60"
      Height          =   195
      Left            =   3240
      TabIndex        =   3
      Top             =   2760
      Width           =   180
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Time:"
      Height          =   195
      Left            =   2760
      TabIndex        =   2
      Top             =   2760
      Width           =   390
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Score:"
      Height          =   195
      Left            =   1080
      TabIndex        =   1
      Top             =   2760
      Width           =   465
   End
   Begin VB.Label lblScore 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   1560
      TabIndex        =   0
      Top             =   2760
      Width           =   90
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   2280
      Shape           =   3  'Circle
      Top             =   1080
      Width           =   135
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   2280
      Shape           =   3  'Circle
      Top             =   1080
      Width           =   135
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   4680
      Y1              =   2520
      Y2              =   2520
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim gtime As Integer
Dim blah As Integer
Dim go As Boolean
Dim score As Integer
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If go = True Then
    Select Case KeyCode
        Case 37
            If Shape1.Left > 0 Then
                Shape1.Left = Shape1.Left - 100
            End If
        Case 38
            If Shape1.Top > 0 Then
                Shape1.Top = Shape1.Top - 100
            End If
        Case 39
            If Shape1.Left < Form1.Width - Shape1.Width Then
                Shape1.Left = Shape1.Left + 100
            End If
        Case 40
            If Shape1.Top < Form1.Height - (Form1.Height - Line1.Y1) - Shape1.Height Then
                Shape1.Top = Shape1.Top + 100
            End If
    End Select
    End If
End Sub

Private Sub Form_Load()
    Dim x As Integer
    Dim y As Integer
    
    Randomize
    y = (Rnd * (2520 - Shape1.Width)) + 1
    x = (Rnd * (4680 - Shape1.Height)) + 1
    Shape1.Left = x
    Shape1.Top = y
    lblx = x
    lbly = y
    
    gtime = 60
    go = True
    
    lblEnd.Visible = False
    Timer1.Enabled = True
    Timer2.Enabled = True
End Sub

Private Sub Timer1_Timer()
    Randomize
    blah = Rnd * 4
    mover
End Sub

Private Sub Timer2_Timer()
    If gtime > 0 Then
        gtime = gtime - 1
        lblTime = gtime
    Else
        Timer1.Enabled = False
        go = False
        lblEnd.Visible = True
    End If
End Sub

Private Sub Timer3_Timer()
    If Shape1.Left < Shape2.Left + Shape2.Width And Shape1.Left > Shape2.Left - Shape2.Width Then
        If Shape1.Top < Shape2.Top + Shape2.Height And Shape1.Top > Shape2.Top - Shape2.Height Then
            score = score + 1
            lblScore = score
            blah = 3
            mover
        End If
    End If
End Sub
Function mover()
    Dim x As Integer
    Dim y As Integer
    
        Select Case blah
            Case 3
                Randomize
                x = (Rnd * (4680 - Shape2.Height)) + 1
                y = (Rnd * (2520 - Shape2.Width)) + 1
                Shape2.Left = x
                Shape2.Top = y
            Case Else
        End Select
End Function
