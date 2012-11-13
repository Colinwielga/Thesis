VERSION 5.00
Begin VB.Form frmHard
   BackColor       =   &H00000000&
   Caption         =   "Hard"
   ClientHeight    =   5055
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   ScaleHeight     =   5055
   ScaleWidth      =   5295
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReset
      Caption         =   "Reset"
      Height          =   495
      Left            =   3840
      TabIndex        =   88
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton cmdSolution
      Caption         =   "Solution"
      Height          =   495
      Left            =   3840
      TabIndex        =   87
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox txt77
      Height          =   285
      Left            =   120
      TabIndex        =   83
      Top             =   3240
      Width           =   255
   End
   Begin VB.TextBox txt78
      Height          =   285
      Left            =   480
      TabIndex        =   82
      Top             =   3240
      Width           =   255
   End
   Begin VB.TextBox txt79
      Height          =   285
      Left            =   840
      TabIndex        =   81
      Top             =   3240
      Width           =   255
   End
   Begin VB.TextBox txt87
      Height          =   285
      Left            =   1320
      TabIndex        =   80
      Top             =   3240
      Width           =   255
   End
   Begin VB.TextBox txt88
      Height          =   285
      Left            =   1680
      TabIndex        =   79
      Top             =   3240
      Width           =   255
   End
   Begin VB.TextBox txt89
      Height          =   285
      Left            =   2040
      TabIndex        =   78
      Top             =   3240
      Width           =   255
   End
   Begin VB.TextBox txt97
      Height          =   285
      Left            =   2520
      TabIndex        =   77
      Top             =   3240
      Width           =   255
   End
   Begin VB.TextBox txt98
      Height          =   285
      Left            =   2880
      TabIndex        =   76
      Top             =   3240
      Width           =   255
   End
   Begin VB.TextBox txt99
      Height          =   285
      Left            =   3240
      TabIndex        =   75
      Top             =   3240
      Width           =   255
   End
   Begin VB.TextBox txt96
      Height          =   285
      Left            =   3240
      TabIndex        =   74
      Top             =   2880
      Width           =   255
   End
   Begin VB.TextBox txt95
      Height          =   285
      Left            =   2880
      TabIndex        =   73
      Top             =   2880
      Width           =   255
   End
   Begin VB.TextBox txt94
      Height          =   285
      Left            =   2520
      TabIndex        =   72
      Top             =   2880
      Width           =   255
   End
   Begin VB.TextBox txt86
      Height          =   285
      Left            =   2040
      TabIndex        =   71
      Top             =   2880
      Width           =   255
   End
   Begin VB.TextBox txt85
      Height          =   285
      Left            =   1680
      TabIndex        =   70
      Top             =   2880
      Width           =   255
   End
   Begin VB.TextBox txt84
      Height          =   285
      Left            =   1320
      TabIndex        =   69
      Top             =   2880
      Width           =   255
   End
   Begin VB.TextBox txt76
      Height          =   285
      Left            =   840
      TabIndex        =   68
      Top             =   2880
      Width           =   255
   End
   Begin VB.TextBox txt75
      Height          =   285
      Left            =   480
      TabIndex        =   67
      Top             =   2880
      Width           =   255
   End
   Begin VB.TextBox txt74
      Height          =   285
      Left            =   120
      TabIndex        =   66
      Top             =   2880
      Width           =   255
   End
   Begin VB.TextBox txt71
      Height          =   285
      Left            =   120
      TabIndex        =   65
      Top             =   2520
      Width           =   255
   End
   Begin VB.TextBox txt72
      Height          =   285
      Left            =   480
      TabIndex        =   64
      Top             =   2520
      Width           =   255
   End
   Begin VB.TextBox txt73
      Height          =   285
      Left            =   840
      TabIndex        =   63
      Top             =   2520
      Width           =   255
   End
   Begin VB.TextBox txt81
      Height          =   285
      Left            =   1320
      TabIndex        =   62
      Top             =   2520
      Width           =   255
   End
   Begin VB.TextBox txt82
      Height          =   285
      Left            =   1680
      TabIndex        =   61
      Top             =   2520
      Width           =   255
   End
   Begin VB.TextBox txt83
      Height          =   285
      Left            =   2040
      TabIndex        =   60
      Top             =   2520
      Width           =   255
   End
   Begin VB.TextBox txt91
      Height          =   285
      Left            =   2520
      TabIndex        =   59
      Top             =   2520
      Width           =   255
   End
   Begin VB.TextBox txt92
      Height          =   285
      Left            =   2880
      TabIndex        =   58
      Top             =   2520
      Width           =   255
   End
   Begin VB.TextBox txt93
      Height          =   285
      Left            =   3240
      TabIndex        =   57
      Top             =   2520
      Width           =   255
   End
   Begin VB.TextBox txt69
      Height          =   285
      Left            =   3240
      TabIndex        =   56
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox txt68
      Height          =   285
      Left            =   2880
      TabIndex        =   55
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox txt67
      Height          =   285
      Left            =   2520
      TabIndex        =   54
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox txt59
      Height          =   285
      Left            =   2040
      TabIndex        =   53
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox txt58
      Height          =   285
      Left            =   1680
      TabIndex        =   52
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox txt57
      Height          =   285
      Left            =   1320
      TabIndex        =   51
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox txt49
      Height          =   285
      Left            =   840
      TabIndex        =   50
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox txt48
      Height          =   285
      Left            =   480
      TabIndex        =   49
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox txt47
      Height          =   285
      Left            =   120
      TabIndex        =   48
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox txt44
      Height          =   285
      Left            =   120
      TabIndex        =   47
      Top             =   1680
      Width           =   255
   End
   Begin VB.TextBox txt45
      Height          =   285
      Left            =   480
      TabIndex        =   46
      Top             =   1680
      Width           =   255
   End
   Begin VB.TextBox txt46
      Height          =   285
      Left            =   840
      TabIndex        =   45
      Top             =   1680
      Width           =   255
   End
   Begin VB.TextBox txt54
      Height          =   285
      Left            =   1320
      TabIndex        =   44
      Top             =   1680
      Width           =   255
   End
   Begin VB.TextBox txt55
      Height          =   285
      Left            =   1680
      TabIndex        =   43
      Top             =   1680
      Width           =   255
   End
   Begin VB.TextBox txt56
      Height          =   285
      Left            =   2040
      TabIndex        =   42
      Top             =   1680
      Width           =   255
   End
   Begin VB.TextBox txt64
      Height          =   285
      Left            =   2520
      TabIndex        =   41
      Top             =   1680
      Width           =   255
   End
   Begin VB.TextBox txt65
      Height          =   285
      Left            =   2880
      TabIndex        =   40
      Top             =   1680
      Width           =   255
   End
   Begin VB.TextBox txt66
      Height          =   285
      Left            =   3240
      TabIndex        =   39
      Top             =   1680
      Width           =   255
   End
   Begin VB.TextBox txt63
      Height          =   285
      Left            =   3240
      TabIndex        =   38
      Top             =   1320
      Width           =   255
   End
   Begin VB.TextBox txt62
      Height          =   285
      Left            =   2880
      TabIndex        =   37
      Top             =   1320
      Width           =   255
   End
   Begin VB.TextBox txt61
      Height          =   285
      Left            =   2520
      TabIndex        =   36
      Top             =   1320
      Width           =   255
   End
   Begin VB.TextBox txt53
      Height          =   285
      Left            =   2040
      TabIndex        =   35
      Top             =   1320
      Width           =   255
   End
   Begin VB.TextBox txt52
      Height          =   285
      Left            =   1680
      TabIndex        =   34
      Top             =   1320
      Width           =   255
   End
   Begin VB.TextBox txt51
      Height          =   285
      Left            =   1320
      TabIndex        =   33
      Top             =   1320
      Width           =   255
   End
   Begin VB.TextBox txt43
      Height          =   285
      Left            =   840
      TabIndex        =   32
      Top             =   1320
      Width           =   255
   End
   Begin VB.TextBox txt42
      Height          =   285
      Left            =   480
      TabIndex        =   31
      Top             =   1320
      Width           =   255
   End
   Begin VB.TextBox txt41
      Height          =   285
      Left            =   120
      TabIndex        =   30
      Top             =   1320
      Width           =   255
   End
   Begin VB.TextBox txt17
      Height          =   285
      Left            =   120
      TabIndex        =   29
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox txt18
      Height          =   285
      Left            =   480
      TabIndex        =   28
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox txt19
      Height          =   285
      Left            =   840
      TabIndex        =   27
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox txt27
      Height          =   285
      Left            =   1320
      TabIndex        =   26
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox txt28
      Height          =   285
      Left            =   1680
      TabIndex        =   25
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox txt29
      Height          =   285
      Left            =   2040
      TabIndex        =   24
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox txt37
      Height          =   285
      Left            =   2520
      TabIndex        =   23
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox txt38
      Height          =   285
      Left            =   2880
      TabIndex        =   22
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox txt39
      Height          =   285
      Left            =   3240
      TabIndex        =   21
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox txt36
      Height          =   285
      Left            =   3240
      TabIndex        =   20
      Top             =   480
      Width           =   255
   End
   Begin VB.TextBox txt35
      Height          =   285
      Left            =   2880
      TabIndex        =   19
      Top             =   480
      Width           =   255
   End
   Begin VB.TextBox txt34
      Height          =   285
      Left            =   2520
      TabIndex        =   18
      Top             =   480
      Width           =   255
   End
   Begin VB.TextBox txt26
      Height          =   285
      Left            =   2040
      TabIndex        =   17
      Top             =   480
      Width           =   255
   End
   Begin VB.TextBox txt25
      Height          =   285
      Left            =   1680
      TabIndex        =   16
      Top             =   480
      Width           =   255
   End
   Begin VB.TextBox txt24
      Height          =   285
      Left            =   1320
      TabIndex        =   15
      Tag             =   "S"
      Top             =   480
      Width           =   255
   End
   Begin VB.TextBox txt33
      Height          =   285
      Left            =   3240
      TabIndex        =   14
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox txt32
      Height          =   285
      Left            =   2880
      TabIndex        =   13
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox txt31
      Height          =   285
      Left            =   2520
      TabIndex        =   12
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox txt23
      Height          =   285
      Left            =   2040
      TabIndex        =   11
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox txt22
      Height          =   285
      Left            =   1680
      TabIndex        =   10
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox txt21
      Height          =   285
      Left            =   1320
      TabIndex        =   9
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox txt16
      Height          =   285
      Left            =   840
      TabIndex        =   8
      Top             =   480
      Width           =   255
   End
   Begin VB.TextBox txt15
      Height          =   285
      Left            =   480
      TabIndex        =   7
      Top             =   480
      Width           =   255
   End
   Begin VB.TextBox txt14
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   255
   End
   Begin VB.TextBox txt13
      Height          =   285
      Left            =   840
      TabIndex        =   5
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox txt12
      Height          =   285
      Left            =   480
      TabIndex        =   4
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox txt11
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   255
   End
   Begin VB.Timer tmrTimer
      Interval        =   100
      Left            =   360
      Top             =   4440
   End
   Begin VB.CommandButton cmdCheck
      Caption         =   "Check"
      Height          =   495
      Left            =   3840
      TabIndex        =   2
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton cmdGoBack
      Caption         =   "Go Back"
      Height          =   495
      Left            =   3840
      TabIndex        =   1
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton cmdQuit
      Caption         =   "Quit"
      Height          =   495
      Left            =   3840
      TabIndex        =   0
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label Label6
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3120
      TabIndex        =   86
      Top             =   3840
      Width           =   375
   End
   Begin VB.Label lblSudoku
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Hard"
      BeginProperty Font
         Name            =   "Pristina"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   1695
      Left            =   -1320
      TabIndex        =   85
      Top             =   3720
      Width           =   6015
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblTimer
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3840
      TabIndex        =   84
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Line Line1
      BorderColor     =   &H008080FF&
      X1              =   1200
      X2              =   1200
      Y1              =   120
      Y2              =   3480
   End
   Begin VB.Line Line2
      BorderColor     =   &H008080FF&
      X1              =   2400
      X2              =   2400
      Y1              =   120
      Y2              =   3480
   End
   Begin VB.Line Line3
      BorderColor     =   &H008080FF&
      X1              =   120
      X2              =   3480
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line4
      BorderColor     =   &H008080FF&
      X1              =   120
      X2              =   3480
      Y1              =   2400
      Y2              =   2400
   End
End
Attribute VB_Name = "frmHard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Found As Boolean
'This checks your answer
Private Sub cmdCheck_Click()
    Dim R As Integer, C As Integer
    Dim cA As Integer, cB As Integer, cC As Integer, cD As Integer, cE As Integer, cF As Integer, cG As Integer, cH As Integer, cI As Integer
    Dim rA As Single, rB As Integer, rC As Integer, rD As Integer, rE As Integer, rF As Integer, rG As Integer, rH As Integer, rI As Integer

    rA = Val(txt11.Text) + Val(txt12.Text) + Val(txt13.Text) + Val(txt21.Text) + Val(txt22.Text) + Val(txt23.Text) + Val(txt31.Text) + Val(txt32.Text) + Val(txt33.Text)
    rB = Val(txt14.Text) + Val(txt15.Text) + Val(txt16.Text) + Val(txt24.Text) + Val(txt25.Text) + Val(txt26.Text) + Val(txt34.Text) + Val(txt35.Text) + Val(txt36.Text)
    rC = Val(txt17.Text) + Val(txt18.Text) + Val(txt19.Text) + Val(txt27.Text) + Val(txt28.Text) + Val(txt29.Text) + Val(txt37.Text) + Val(txt38.Text) + Val(txt39.Text)
    rD = Val(txt41.Text) + Val(txt42.Text) + Val(txt43.Text) + Val(txt51.Text) + Val(txt52.Text) + Val(txt53.Text) + Val(txt61.Text) + Val(txt62.Text) + Val(txt63.Text)
    rE = Val(txt44.Text) + Val(txt45.Text) + Val(txt46.Text) + Val(txt54.Text) + Val(txt55.Text) + Val(txt56.Text) + Val(txt64.Text) + Val(txt65.Text) + Val(txt66.Text)
    rF = Val(txt47.Text) + Val(txt48.Text) + Val(txt49.Text) + Val(txt57.Text) + Val(txt58.Text) + Val(txt59.Text) + Val(txt67.Text) + Val(txt68.Text) + Val(txt69.Text)
    rG = Val(txt71.Text) + Val(txt72.Text) + Val(txt73.Text) + Val(txt81.Text) + Val(txt82.Text) + Val(txt83.Text) + Val(txt91.Text) + Val(txt92.Text) + Val(txt93.Text)
    rH = Val(txt74.Text) + Val(txt75.Text) + Val(txt76.Text) + Val(txt84.Text) + Val(txt85.Text) + Val(txt86.Text) + Val(txt94.Text) + Val(txt95.Text) + Val(txt96.Text)
    rI = Val(txt77.Text) + Val(txt78.Text) + Val(txt79.Text) + Val(txt87.Text) + Val(txt88.Text) + Val(txt89.Text) + Val(txt97.Text) + Val(txt98.Text) + Val(txt99.Text)
    R = rA + rB + rC + rD + rE + rF + rG + rH + rI
    cA = Val(txt11.Text) + Val(txt14.Text) + Val(txt17.Text) + Val(txt21.Text) + Val(txt24.Text) + Val(txt27.Text) + Val(txt31.Text) + Val(txt34.Text) + Val(txt37.Text)
    cB = Val(txt12.Text) + Val(txt15.Text) + Val(txt18.Text) + Val(txt22.Text) + Val(txt25.Text) + Val(txt28.Text) + Val(txt32.Text) + Val(txt35.Text) + Val(txt38.Text)
    cC = Val(txt13.Text) + Val(txt16.Text) + Val(txt19.Text) + Val(txt23.Text) + Val(txt26.Text) + Val(txt29.Text) + Val(txt33.Text) + Val(txt36.Text) + Val(txt39.Text)
    cD = Val(txt41.Text) + Val(txt44.Text) + Val(txt47.Text) + Val(txt51.Text) + Val(txt54.Text) + Val(txt57.Text) + Val(txt61.Text) + Val(txt64.Text) + Val(txt67.Text)
    cE = Val(txt42.Text) + Val(txt45.Text) + Val(txt48.Text) + Val(txt52.Text) + Val(txt55.Text) + Val(txt58.Text) + Val(txt62.Text) + Val(txt65.Text) + Val(txt68.Text)
    cF = Val(txt43.Text) + Val(txt46.Text) + Val(txt49.Text) + Val(txt53.Text) + Val(txt56.Text) + Val(txt59.Text) + Val(txt63.Text) + Val(txt66.Text) + Val(txt69.Text)
    cG = Val(txt71.Text) + Val(txt74.Text) + Val(txt77.Text) + Val(txt81.Text) + Val(txt84.Text) + Val(txt87.Text) + Val(txt91.Text) + Val(txt94.Text) + Val(txt97.Text)
    cH = Val(txt72.Text) + Val(txt75.Text) + Val(txt78.Text) + Val(txt82.Text) + Val(txt85.Text) + Val(txt88.Text) + Val(txt92.Text) + Val(txt95.Text) + Val(txt98.Text)
    cI = Val(txt73.Text) + Val(txt76.Text) + Val(txt79.Text) + Val(txt83.Text) + Val(txt86.Text) + Val(txt89.Text) + Val(txt93.Text) + Val(txt96.Text) + Val(txt99.Text)
    C = cA + cB + cC + cD + cE + cF + cG + cH + cI
    HardTime = lblTimer
    Found = False
    If R = 405 And C = 405 Then
        frmHappy.Show
        MsgBox "Congratulations " & FirstName & ", you completed the Hard puzzle in " & HardTime & "seconds!", , "WINNER!"
        Found = True
        HardName = LastName & ", " & FirstName
    Else
        MsgBox "Sorry, that is not the solution.", , "Try Again"
    End If
    If Found = True Then
        frmHappy.Hide
        frmHard.Hide
        frmSudoku.Show
    End If
End Sub

'This brings you back to the main page
Private Sub cmdGoBack_Click()
    frmSudoku.Show
    frmMedium.Hide
End Sub

'This quits the program
Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdReset_Click()
    lblTimer = 0
    lblTimer = lblTimer + 0.1
    txt11.Text = ""
    txt12.Text = ""
    txt13.Text = ""
    txt15.Text = ""
    txt16.Text = ""
    txt19.Text = ""
    txt21.Text = ""
    txt22.Text = ""
    txt23.Text = ""
    txt25.Text = ""
    txt26.Text = ""
    txt27.Text = ""
    txt29.Text = ""
    txt32.Text = ""
    txt34.Text = ""
    txt35.Text = ""
    txt36.Text = ""
    txt37.Text = ""
    txt38.Text = ""
    txt39.Text = ""
    txt41.Text = ""
    txt42.Text = ""
    txt43.Text = ""
    txt44.Text = ""
    txt46.Text = ""
    txt49.Text = ""
    txt51.Text = ""
    txt53.Text = ""
    txt55.Text = ""
    txt57.Text = ""
    txt59.Text = ""
    txt61.Text = ""
    txt64.Text = ""
    txt66.Text = ""
    txt67.Text = ""
    txt68.Text = ""
    txt69.Text = ""
    txt71.Text = ""
    txt72.Text = ""
    txt73.Text = ""
    txt74.Text = ""
    txt75.Text = ""
    txt76.Text = ""
    txt78.Text = ""
    txt81.Text = ""
    txt83.Text = ""
    txt84.Text = ""
    txt85.Text = ""
    txt87.Text = ""
    txt88.Text = ""
    txt89.Text = ""
    txt91.Text = ""
    txt94.Text = ""
    txt95.Text = ""
    txt97.Text = ""
    txt98.Text = ""
    txt99.Text = ""
End Sub

'This brings you to the solutions page
Private Sub cmdSolution_Click()
    frmSolution.Show
End Sub

'This loads your puzzle
Private Sub Form_Load()
    MsgBox "Enjoy your Hard puzzle " & FirstName & ". Click Ok to begin.", , "Hard"
    txt14.Enabled = False
    txt17.Enabled = False
    txt18.Enabled = False
    txt24.Enabled = False
    txt28.Enabled = False
    txt31.Enabled = False
    txt33.Enabled = False
    txt45.Enabled = False
    txt47.Enabled = False
    txt48.Enabled = False
    txt52.Enabled = False
    txt54.Enabled = False
    txt56.Enabled = False
    txt58.Enabled = False
    txt62.Enabled = False
    txt63.Enabled = False
    txt65.Enabled = False
    txt77.Enabled = False
    txt79.Enabled = False
    txt82.Enabled = False
    txt86.Enabled = False
    txt92.Enabled = False
    txt93.Enabled = False
    txt96.Enabled = False
    txt14.Text = 9
    txt17.Text = 1
    txt18.Text = 4
    txt24.Text = 1
    txt28.Text = 5
    txt31.Text = 6
    txt33.Text = 4
    txt45.Text = 7
    txt47.Text = 5
    txt48.Text = 9
    txt52.Text = 1
    txt54.Text = 3
    txt56.Text = 5
    txt58.Text = 2
    txt62.Text = 2
    txt63.Text = 7
    txt65.Text = 8
    txt77.Text = 8
    txt79.Text = 3
    txt82.Text = 8
    txt86.Text = 3
    txt92.Text = 9
    txt93.Text = 5
    txt96.Text = 3
End Sub

'This is your timer
Private Sub tmrTimer_Timer()
    Found = False

    If Found = False Then
        lblTimer.Caption = lblTimer.Caption + 0.1
    Else
        lblTimer.Caption = lblTimer.Caption
    End If
End Sub


