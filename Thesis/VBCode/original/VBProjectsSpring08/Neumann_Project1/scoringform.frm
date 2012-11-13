VERSION 5.00
Begin VB.Form scoringform 
   Caption         =   "Form1"
   ClientHeight    =   7980
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14415
   LinkTopic       =   "Form1"
   Picture         =   "scoringform.frx":0000
   ScaleHeight     =   7980
   ScaleWidth      =   14415
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Reset game"
      Height          =   615
      Left            =   11400
      TabIndex        =   45
      Top             =   3120
      Width           =   1695
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "back home"
      Height          =   615
      Left            =   11400
      TabIndex        =   42
      Top             =   2280
      Width           =   1695
   End
   Begin VB.PictureBox picten 
      Height          =   495
      Left            =   12360
      ScaleHeight     =   435
      ScaleWidth      =   1275
      TabIndex        =   41
      Top             =   1200
      Width           =   1335
   End
   Begin VB.PictureBox picnine 
      Height          =   495
      Left            =   10920
      ScaleHeight     =   435
      ScaleWidth      =   915
      TabIndex        =   40
      Top             =   1200
      Width           =   975
   End
   Begin VB.PictureBox piceight 
      Height          =   495
      Left            =   9600
      ScaleHeight     =   435
      ScaleWidth      =   915
      TabIndex        =   39
      Top             =   1200
      Width           =   975
   End
   Begin VB.PictureBox picseven 
      Height          =   495
      Left            =   8280
      ScaleHeight     =   435
      ScaleWidth      =   915
      TabIndex        =   38
      Top             =   1200
      Width           =   975
   End
   Begin VB.PictureBox picsix 
      Height          =   495
      Left            =   6960
      ScaleHeight     =   435
      ScaleWidth      =   915
      TabIndex        =   37
      Top             =   1200
      Width           =   975
   End
   Begin VB.PictureBox picfive 
      Height          =   495
      Left            =   5640
      ScaleHeight     =   435
      ScaleWidth      =   915
      TabIndex        =   36
      Top             =   1200
      Width           =   975
   End
   Begin VB.PictureBox picfour 
      Height          =   495
      Left            =   4320
      ScaleHeight     =   435
      ScaleWidth      =   915
      TabIndex        =   35
      Top             =   1200
      Width           =   975
   End
   Begin VB.PictureBox picthree 
      Height          =   495
      Left            =   3000
      ScaleHeight     =   435
      ScaleWidth      =   915
      TabIndex        =   34
      Top             =   1200
      Width           =   975
   End
   Begin VB.PictureBox pictwo 
      Height          =   495
      Left            =   1680
      ScaleHeight     =   435
      ScaleWidth      =   915
      TabIndex        =   33
      Top             =   1200
      Width           =   975
   End
   Begin VB.PictureBox picone 
      Height          =   495
      Left            =   360
      ScaleHeight     =   435
      ScaleWidth      =   915
      TabIndex        =   32
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton cmdcompute 
      Caption         =   "Compute your score"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6480
      TabIndex        =   31
      Top             =   2280
      Width           =   2055
   End
   Begin VB.TextBox txtnineteen 
      Height          =   495
      Left            =   12120
      TabIndex        =   20
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox txttwenty 
      Height          =   495
      Left            =   12720
      TabIndex        =   19
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox txttwentyone 
      Height          =   495
      Left            =   13320
      TabIndex        =   18
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox txtten 
      Height          =   495
      Left            =   6120
      TabIndex        =   17
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox txteleven 
      Height          =   495
      Left            =   6840
      TabIndex        =   16
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox txttwelve 
      Height          =   495
      Left            =   7440
      TabIndex        =   15
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox txtthirteen 
      Height          =   495
      Left            =   8160
      TabIndex        =   14
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox txtfourteen 
      Height          =   495
      Left            =   8760
      TabIndex        =   13
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox txtfifeteen 
      Height          =   495
      Left            =   9480
      TabIndex        =   12
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox txtsixteen 
      Height          =   495
      Left            =   10080
      TabIndex        =   11
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox txtseventeen 
      Height          =   495
      Left            =   10800
      TabIndex        =   10
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox txteighteen 
      Height          =   495
      Left            =   11400
      TabIndex        =   9
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox txttwo 
      Height          =   495
      Left            =   840
      TabIndex        =   8
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox txtthree 
      Height          =   495
      Left            =   1560
      TabIndex        =   7
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox txtfour 
      Height          =   495
      Left            =   2160
      TabIndex        =   6
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox txtfive 
      Height          =   495
      Left            =   2880
      TabIndex        =   5
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox txtsix 
      Height          =   495
      Left            =   3480
      TabIndex        =   4
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox txtseven 
      Height          =   495
      Left            =   4200
      TabIndex        =   3
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox txteight 
      Height          =   495
      Left            =   4800
      TabIndex        =   2
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox txtnine 
      Height          =   495
      Left            =   5520
      TabIndex        =   1
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox txtone 
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label12 
      Caption         =   "If you get a strike enter 0 in for the next ball, and enter 0 for the 12th ball if there is no strike or spare in the 10th frame "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   44
      Top             =   1920
      Width           =   5175
   End
   Begin VB.Label Label11 
      Caption         =   $"scoringform.frx":37F8F
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   360
      TabIndex        =   43
      Top             =   2760
      Width           =   4575
   End
   Begin VB.Label Label10 
      Caption         =   "10th frame"
      Height          =   375
      Left            =   12600
      TabIndex        =   30
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label9 
      Caption         =   "9th frame"
      Height          =   375
      Left            =   11040
      TabIndex        =   29
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label8 
      Caption         =   "8th frame"
      Height          =   375
      Left            =   9720
      TabIndex        =   28
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label7 
      Caption         =   "7th frame"
      Height          =   375
      Left            =   8400
      TabIndex        =   27
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "6th frame"
      Height          =   375
      Left            =   7080
      TabIndex        =   26
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "5th frame"
      Height          =   375
      Left            =   5760
      TabIndex        =   25
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "4th frame"
      Height          =   375
      Left            =   4440
      TabIndex        =   24
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "3rd frame"
      Height          =   375
      Left            =   3120
      TabIndex        =   23
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "2nd frame"
      Height          =   375
      Left            =   1800
      TabIndex        =   22
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "1st frame"
      Height          =   375
      Left            =   480
      TabIndex        =   21
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "scoringform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Bowling prodject
'scoring form
'Zach Neumann
'3/30/2008
'This form teaches a person how to score with a interactive score board that allows
'them to enter in scores for each ball. Also can be used for a bowler to record the game that they are bowling

Option Explicit
Dim score As Integer, ballone As Integer, balltwo As Integer, ballthree As Integer, ballfour As Integer, ballfive As Integer, ballsix As Integer, ballseven As Integer, balleight As Integer, ballnine As Integer, ballten As Integer, balleleven As Integer, balltwelve As Integer, ballthirteen As Integer, ballfourteen As Integer, ballfifeteen As Integer, ballsixteen As Integer, ballseventeen As Integer, balleighteen As Integer, ballnineteen As Integer, balltwenty As Integer, balltwentyone As Integer

Private Sub cmdback_Click()
    scoringform.Hide
    startform.Show
    
End Sub

Private Sub cmdcompute_Click()
Dim total As Integer
'defines all the text boxes
ballone = txtone.Text
balltwo = txttwo.Text
ballthree = txtthree.Text
ballfour = txtfour.Text
ballfive = txtfive.Text
ballsix = txtsix.Text
ballseven = txtseven.Text
balleight = txteight.Text
ballnine = txtnine.Text
ballten = txtten.Text
balleleven = txteleven.Text
balltwelve = txttwelve.Text
ballthirteen = txtthirteen.Text
ballfourteen = txtfourteen.Text
ballfifeteen = txtfifeteen.Text
ballsixteen = txtsixteen.Text
ballseventeen = txtseventeen.Text
balleighteen = txteighteen.Text
ballnineteen = txtnineteen.Text
balltwenty = txttwenty.Text
balltwentyone = txttwentyone.Text
'all the following code is used for scoring, it takes into account what must be
'done if the bowler gets a strike and a spare and also the 3 ball 10th frame
If ballone + balltwo > 10 Then
    MsgBox "Please enter a total of 10 pins at the most", , "Error!"
End If
    
If ballone = 10 Then
    balltwo = 0
    picone.Print "X"
    If ballthree = 10 Then
        picone.Print score + ballthree + ballfive + ballone
        score = score + ballthree + ballfive + ballone
    Else
        picone.Print score + ballthree + ballfour + ballone
        score = score + ballthree + ballfour + ballone
    End If
ElseIf ballone + balltwo = 10 Then
    picone.Print "/"
    picone.Print score + ballthree + ballthree + balltwo
    score = score + ballthree + ballone + balltwo
ElseIf (ballone + balltwo) < 10 Then
    picone.Print ballone + balltwo + score
    score = score + ballone + balltwo
End If

If ballthree + ballfour > 10 Then
    MsgBox "Please enter a total of 10 pins at the most", , "Error!"
End If
If ballthree = 10 Then
    ballfour = 0
    pictwo.Print "X"
    If ballfive = 10 Then
        pictwo.Print score + ballfive + ballseven + ballthree
        score = score + ballfive + ballseven + ballthree
    Else
        pictwo.Print score + ballfive + ballsix
        score = score + ballfive + ballsix + ballthree
    End If
ElseIf ballthree + ballfour = 10 Then
    pictwo.Print "/"
    pictwo.Print score + ballfive
    score = score + ballfive + ballthree + ballfour
ElseIf ballthree + ballfour < 10 Then
    pictwo.Print score + ballthree + ballfour
    score = score + ballthree + ballfour
End If

If ballfive + ballsix > 10 Then
    MsgBox "Please enter a total of 10 pins at the most", , "Error!"
End If
If ballfive = 10 Then
    ballsix = 0
    picthree.Print "X"
    If ballseven = 10 Then
        picthree.Print score + ballseven + ballnine + ballfive
        score = score + ballseven + ballnine + ballfive
    Else
        picthree.Print score + ballseven + balleight + ballfive
        score = score + ballseven + balleight + ballfive
    End If
ElseIf ballfive + ballsix = 10 Then
    picthree.Print "/"
    picthree.Print score + ballseven + ballfive + ballsix
    score = score + ballseven + ballfive + ballsix
ElseIf ballfive + ballsix < 10 Then
    picthree.Print score + ballfive + ballsix
    score = score + ballfive + ballsix
End If

If ballseven + balleight > 10 Then
    MsgBox "Please enter a total of 10 pins at the most", , "Error!"
End If
If ballseven = 10 Then
    balleight = 0
    picfour.Print "X"
    If ballnine = 10 Then
        picfour.Print score + ballseven + ballnine + balleleven
        score = score + ballseven + ballnine + balleleven
    Else
        picfour.Print score + ballnine + ballten
        score = score + ballnine + ballten
    End If
ElseIf ballseven + balleight = 10 Then
    picfour.Print "/"
    picfour.Print score + ballnine + ballseven + balleight
    score = score + ballnine + ballseven + balleight
ElseIf ballseven + balleight < 10 Then
    picfour.Print score + ballseven + balleight
    score = score + ballseven + balleight
End If

If ballnine + ballten > 10 Then
    MsgBox "Please enter a total of 10 pins at the most", , "Error!"
End If
If ballnine = 10 Then
    ballten = 0
    picfive.Print "X"
    If balleleven = 10 Then
        picfive.Print score + balleleven + ballthirteen + ballnine
        score = score + balleleven + ballthirteen + ballnine
    Else
        picfive.Print score + balleleven + balltwelve
        score = score + balleleven + balltwelve + ballnine
    End If
ElseIf ballnine + ballten = 10 Then
    picfive.Print "/"
    picfive.Print score + ballnine + ballten + balleleven
    score = score + ballnine + ballten + balleleven
ElseIf ballnine + ballten < 10 Then
    picfive.Print score + ballnine + ballten
    score = score + ballnine + ballten
End If
    
If balleleven + balltwelve > 10 Then
    MsgBox "Please enter a total of 10 pins at the most", , "Error!"
End If
If balleleven = 10 Then
    balltwelve = 0
    picsix.Print "X"
    If ballthirteen = 10 Then
        picsix.Print score + balleleven + ballthirteen + ballfifeteen
        score = score + balleleven + ballthirteen + ballfifeteen
    Else
        picsix.Print score + ballthirteen + ballfourteen + balleleven
        score = score + ballfourteen + ballthirteen + balleleven
    End If
ElseIf balleleven + balltwelve = 10 Then
    picsix.Print "/"
    picsix.Print score + balleleven + balltwelve + ballthirteen
    score = score + balleleven + balltwelve + ballthirteen
ElseIf balleleven + balltwelve < 10 Then
    picsix.Print score + balleleven + balltwelve
    score = score + balleleven + balltwelve
End If

If ballthirteen + ballfourteen > 10 Then
    MsgBox "Please enter a total of 10 pins at the most", , "Error!"
End If
If ballthirteen = 10 Then
    ballfourteen = 0
    picseven.Print "X"
    If ballfifeteen = 10 Then
        picseven.Print score + ballseventeen + ballthirteen + ballfifeteen
        score = score + ballseventeen + ballthirteen + ballfifeteen
    Else
        picseven.Print score + ballthirteen + ballsixteen + ballfifeteen
        score = score + ballsixteen + ballthirteen + ballfifeteen
    End If
ElseIf ballthirteen + ballfourteen = 10 Then
    picseven.Print "/"
    picseven.Print score + ballfifeteen + ballfourteen + ballthirteen
    score = score + ballfifeteen + ballfourteen + ballthirteen
ElseIf ballthirteen + ballfourteen < 10 Then
    picseven.Print score + balleleven + balltwelve
    score = score + balleleven + balltwelve
End If

If ballfifeteen + ballsixteen > 10 Then
    MsgBox "Please enter a total of 10 pins at the most", , "Error!"
End If
If ballfifeteen = 10 Then
    ballsixteen = 0
    piceight.Print "X"
    If ballseventeen = 10 Then
        piceight.Print score + ballseventeen + ballnineteen + ballfifeteen
        score = score + ballseventeen + ballthirteen + ballfifeteen
    Else
        piceight.Print score + balleighteen + ballseventeen + ballfifeteen
        score = score + balleighteen + ballseventeen + ballfifeteen
    End If
ElseIf ballfifeteen + ballsixteen = 10 Then
    piceight.Print "/"
    piceight.Print score + ballfifeteen + ballsixteen + ballseventeen
    score = score + ballfifeteen + ballsixteen + ballseventeen
ElseIf ballfifeteen + ballsixteen < 10 Then
    piceight.Print score + ballfifeteen + ballsixteen
    score = score + ballfifeteen + ballsixteen
End If

If ballseventeen + balleighteen > 10 Then
    MsgBox "Please enter a total of 10 pins at the most", , "Error!"
End If
If ballseventeen = 10 Then
    balleighteen = 0
    picnine.Print "X"
    picnine.Print score + ballseventeen + ballnineteen + balltwenty
    score = score + ballseventeen + ballnineteen + balltwenty
ElseIf ballseventeen + balleighteen = 10 Then
    picnine.Print "/"
    picnine.Print score + ballnineteen + balleighteen + ballseventeen
    score = score + ballnineteen + balleighteen + ballseventeen
ElseIf ballseventeen + balleighteen < 10 Then
    picnine.Print score + ballseventeen + balleighteen
    score = score + ballseventeen + balleighteen
End If

If ballnineteen = 10 And balltwenty = 10 And balltwentyone = 10 Then
    picten.Print "XXX"
    picten.Print score + ballnineteen + balltwenty + balltwentyone
ElseIf ballnineteen = 10 And balltwenty = 10 Then
    picten.Print "XX"
    picten.Print score + ballnineteen + balltwenty + balltwentyone
ElseIf ballnineteen + balltwenty = 10 Then
    picten.Print "/"
    picten.Print score + ballnineteen + balltwenty + balltwentyone
ElseIf ballnineteen + balltwenty < 10 Then
    picten.Print score + ballnineteen + balltwenty + balltwentyone
End If

    
End Sub



Private Sub Command1_Click()
picone.Cls
pictwo.Cls
picthree.Cls
picfour.Cls
picfive.Cls
picsix.Cls
picseven.Cls
piceight.Cls
picnine.Cls
picten.Cls
End Sub
