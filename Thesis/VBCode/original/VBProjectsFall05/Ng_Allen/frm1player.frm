VERSION 5.00
Begin VB.Form frm1player 
   Caption         =   "Speed Type - ER"
   ClientHeight    =   6495
   ClientLeft      =   2970
   ClientTop       =   2850
   ClientWidth     =   9030
   Picture         =   "frm1player.frx":0000
   ScaleHeight     =   6495
   ScaleWidth      =   9030
   Begin VB.CommandButton cmdstop 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Stop the Game!"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton cmdview 
      BackColor       =   &H00FFFFFF&
      Caption         =   "View High Scores"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4320
      Width           =   975
   End
   Begin VB.OptionButton Optionhard 
      BackColor       =   &H000000FF&
      Caption         =   "Hard"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      TabIndex        =   13
      Top             =   5040
      Width           =   1335
   End
   Begin VB.OptionButton Optionmedium 
      BackColor       =   &H0000FFFF&
      Caption         =   "Medium"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      TabIndex        =   12
      Top             =   4680
      Width           =   1335
   End
   Begin VB.OptionButton Optioneasy 
      BackColor       =   &H0000FF00&
      Caption         =   "Easy"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      TabIndex        =   11
      Top             =   4320
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.CommandButton cmdmain 
      Caption         =   "Go Back To Main Menu"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   10
      Top             =   5160
      Width           =   975
   End
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit :("
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   6000
      Width           =   975
   End
   Begin VB.CommandButton cmdenter 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enter In Word"
      Default         =   -1  'True
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4320
      Width           =   1335
   End
   Begin VB.TextBox txtinput 
      Height          =   975
      Left            =   3600
      TabIndex        =   7
      Top             =   4320
      Width           =   1815
   End
   Begin VB.PictureBox picwordbox 
      Height          =   975
      Left            =   1320
      ScaleHeight     =   915
      ScaleWidth      =   1995
      TabIndex        =   5
      Top             =   4320
      Width           =   2055
   End
   Begin VB.CommandButton cmdGo 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Lets play!"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5760
      Width           =   2295
   End
   Begin VB.PictureBox picOne 
      Height          =   765
      Left            =   7080
      Picture         =   "frm1player.frx":180D9
      ScaleHeight     =   705
      ScaleWidth      =   1455
      TabIndex        =   1
      Top             =   3120
      Width           =   1515
   End
   Begin VB.PictureBox picCPU 
      Height          =   1005
      Left            =   7080
      Picture         =   "frm1player.frx":1C947
      ScaleHeight     =   945
      ScaleWidth      =   1455
      TabIndex        =   0
      Top             =   1200
      Width           =   1515
   End
   Begin VB.PictureBox picracetrack 
      Height          =   3255
      Left            =   960
      Picture         =   "frm1player.frx":234FF
      ScaleHeight     =   3195
      ScaleWidth      =   7635
      TabIndex        =   2
      Top             =   960
      Width           =   7695
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1320
      Top             =   1560
   End
   Begin VB.Label lblfinish 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Finish Line"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   600
      Width           =   1455
   End
   Begin VB.Line linefinish 
      BorderWidth     =   3
      X1              =   840
      X2              =   840
      Y1              =   960
      Y2              =   4200
   End
   Begin VB.Label lbldirections 
      BackColor       =   &H00000000&
      Caption         =   $"frm1player.frx":28304
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   1320
      TabIndex        =   6
      Top             =   5400
      Width           =   3255
   End
   Begin VB.Label lbltitle 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Good Luck!!"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   4
      Top             =   360
      Width           =   3375
   End
End
Attribute VB_Name = "frm1player"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'CSCI 130 Games to play in class (VBProject.vbp)
'Speed Type - ER(frm1player.frm)
'Done by Allen Ng
'30 October 2005
'This form is the actual game.  I used a random number generator,
'a hall of fame array, a word bank, a timer, and many picture
'boxes to make this game possible.  I also used a module
'to declare public variables so I wouldn't have to declare them
'each time.

Private Sub cmdenter_Click()
    If txtinput.Text = Randomword(RandomNumber) Then
        picOne.Left = picOne.Left - 1000
        txtinput.SetFocus 'brings cursor to textbox
        I = I + 1
    End If
    txtinput.Text = ""
    picwordbox.Cls
    'Generates a Random Number BETWEEN the LOWER and UPPER values
    Randomize
    RandomNumber = Int((310 - 1 + 1) * Rnd + 1)

    picwordbox.Print Randomword(RandomNumber)
End Sub

Private Sub cmdGo_Click()
    cmdenter.Enabled = True
    picwordbox.Cls
    picCPU.Left = 6600 'resets cars back to starting position
    picOne.Left = 6600
    score = 500 'resets score
    Timer1.Enabled = True
    Open App.Path & "\Word Bank.txt" For Input As #10
    For I = 1 To 310
        Input #10, Randomword(I)
    Next I
    Close #10
    'Generates a Random Number BETWEEN the LOWER and UPPER values
    Randomize
    RandomNumber = Int((310 - 1 + 1) * Rnd + 1)
    picwordbox.Print Randomword(RandomNumber)
    txtinput.SetFocus
End Sub

Private Sub cmdmain_Click()
    frm1player.Visible = False
    frmMainmenu.Visible = True
End Sub

Private Sub cmdquit_Click()
    End
End Sub

Private Sub cmdstop_Click()
    Timer1.Enabled = False
    cmdenter.Enabled = False
    txtinput.Text = ""
    picCPU.Left = 6600 'resets cars back to starting position
    picOne.Left = 6600
    picwordbox.Cls
End Sub

Private Sub cmdview_Click()
    frmhighscores.Visible = True
    frm1player.Visible = False
End Sub

Private Sub Timer1_Timer()
    'difficulty corresponding to how fast the car moves
    If Optioneasy.Value = True Then
        picCPU.Left = picCPU.Left - 5
    ElseIf Optionmedium.Value = True Then
        picCPU.Left = picCPU.Left - 50
    ElseIf Optionhard.Value = True Then
        picCPU.Left = picCPU.Left - 100
    End If
    score = score - 1
    If score < 0 Then
        score = 0
    End If
    If picCPU.Left <= linefinish.X1 Or picOne.Left <= linefinish.X1 Then 'once one finishes then the timer shuts down so they stop moving.
        Timer1.Enabled = False
        cmdenter.Enabled = False
            If picCPU.Left < picOne.Left Then
                MsgBox "You Lose!", , "You lose, try again!"
                score = 0
            Else
                MsgBox "You Win!", , "Play Again!"
            End If
    Open App.Path & "\Hall of Fame.Txt" For Input As #1
    For I = 1 To 10
        Input #1, Halloffamename(I), Halloffamescore(I)
    Next I
    Close #1
    Notfound = True
    I = 0
    Do While (Notfound = True And I <= 10)
        I = I + 1
    If score > Halloffamescore(I) Then
        MsgBox score, , "Congratulations You have made it to the Hall of Fame!"
        halloffamenametemp = InputBox("Congratulations on the win.  Please enter in your name to be put into the Hall of Fame!", "Enter in your name.")
        halloffamescoretemp = score
        Notfound = False
    End If
    Loop
    If I = 1 Then
        Halloffamename(I) = halloffamenametemp
        Halloffamescore(I) = halloffamescoretemp
    Else
    For J = 10 To I Step -1 'this should move all scores and names down one slot
                Halloffamename(J) = Halloffamename(J - 1)
                Halloffamescore(J) = Halloffamescore(J - 1)
            Next J
            Halloffamename(I) = halloffamenametemp
            Halloffamescore(I) = halloffamescoretemp
    End If
    End If
    If Timer1.Enabled = False Then
        txtinput.Text = ""
        picCPU.Left = 6600 'resets cars back to starting position
        picOne.Left = 6600
        picwordbox.Cls
    'rewrites the hall of fame listing if you make it on the list
    Open App.Path & "\Hall of Fame.Txt" For Output As #1
    For I = 1 To 10
        Write #1, Halloffamename(I), Halloffamescore(I)
    Next I
    Close #1
    End If

End Sub

