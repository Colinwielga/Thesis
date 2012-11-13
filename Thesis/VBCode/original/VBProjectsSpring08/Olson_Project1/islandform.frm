VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   10905
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form3"
   Picture         =   "island form.frx":0000
   ScaleHeight     =   10905
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdmainpage 
      Caption         =   "Go back to Main Page"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   480
      TabIndex        =   6
      Top             =   7080
      Width           =   2415
   End
   Begin VB.PictureBox picisland 
      AutoSize        =   -1  'True
      Height          =   3255
      Left            =   3840
      ScaleHeight     =   3195
      ScaleWidth      =   4275
      TabIndex        =   5
      Top             =   6600
      Width           =   4335
   End
   Begin VB.CommandButton cmdenter 
      Caption         =   "Enter!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10680
      TabIndex        =   3
      Top             =   3840
      Width           =   2775
   End
   Begin VB.TextBox txtisland 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9840
      TabIndex        =   1
      Top             =   2760
      Width           =   5055
   End
   Begin VB.Label lbloutput 
      BorderStyle     =   1  'Fixed Single
      Height          =   3855
      Left            =   9480
      TabIndex        =   4
      Top             =   6000
      Width           =   5295
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Type any place name seen on map:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9960
      TabIndex        =   2
      Top             =   2280
      Width           =   4095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "What part of the island would you like to go?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   9840
      TabIndex        =   0
      Top             =   240
      Width           =   4695
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project name: Gilligan's Island
'Form name:  island form
'Author:  Emily Olson
'Date written:  March 30, 2008
'Form Objective: inform user about the different areas of the island on "Gilligan's Island

'declare file variables
Dim SSMinnow As String
Dim Turtlerace As String
Dim Tongo As String
Dim Goldmine As String
Dim Hills As String
Dim Lagoon As String
Dim Beach As String
Dim Spider As String
Dim Totempole As String
Dim Quicksand As String
Dim Bowling As String
Dim Puttinggreen As String
Dim Volcano As String
Dim Hut As String

Private Sub cmdenter_Click()
'clear data
    lbloutput = ""
'if-elseif statements that will display data and picture from file according to what is inputed by user

    If LCase(Trim(txtisland.Text)) = LCase("SS Minnow") Then
        Open App.Path & "\ssminnow.txt" For Input As #1
        Do Until EOF(1)
            Input #1, SSMinnow
            lbloutput = lbloutput + SSMinnow
        Loop
        Close #1
        picisland.Picture = LoadPicture("ssminnow.jpg")
    ElseIf LCase(Trim(txtisland.Text)) = LCase("Turtle Races") Then
        Open App.Path & "\turtlerace.txt" For Input As #1
        Do Until EOF(1)
            Input #1, Turtlerace
            lbloutput = lbloutput + Turtlerace
        Loop
        Close #1
        picisland.Picture = LoadPicture("turtle.jpg")
    ElseIf LCase(Trim(txtisland.Text)) = LCase("Tongo") Then
        Open App.Path & "\tongo.txt" For Input As #1
        Do Until EOF(1)
            Input #1, Tongo
            lbloutput = lbloutput + Tongo
        Loop
        Close #1
        picisland.Picture = LoadPicture("tongo.jpg")
    ElseIf LCase(Trim(txtisland.Text)) = LCase("Gold Mine") Then
        Open App.Path & "\goldmine.txt" For Input As #1
        Do Until EOF(1)
            Input #1, Goldmine
            lbloutput = lbloutput + Goldmine
        Loop
        Close #1
        picisland.Picture = LoadPicture("goldmine.jpg")
    ElseIf LCase(Trim(txtisland.Text)) = LCase("Howell's putting green") Then
        Open App.Path & "\puttinggreen.txt" For Input As #1
        Do Until EOF(1)
            Input #1, Puttinggreen
            lbloutput = lbloutput + Puttinggreen
        Loop
        Close #1
        picisland.Picture = LoadPicture("puttinggreen.jpg")
    ElseIf LCase(Trim(txtisland.Text)) = LCase("Howell's Hills") Then
        Open App.Path & "\hills.txt" For Input As #1
        Do Until EOF(1)
            Input #1, Hills
            lbloutput = lbloutput + Hills
        Loop
        Close #1
        picisland.Picture = LoadPicture("hills.jpg")
    ElseIf LCase(Trim(txtisland.Text)) = LCase("Lagoon") Then
        Open App.Path & "\lagoon.txt" For Input As #1
        Do Until EOF(1)
            Input #1, Lagoon
            lbloutput = lbloutput + Lagoon
        Loop
        Close #1
        picisland.Picture = LoadPicture("lagoon.jpg")
    ElseIf LCase(Trim(txtisland.Text)) = LCase("Beach") Then
        Open App.Path & "\beach.txt" For Input As #1
        Do Until EOF(1)
            Input #1, Beach
            lbloutput = lbloutput + Beach
        Loop
        Close #1
        picisland.Picture = LoadPicture("beach.jpg")
    ElseIf LCase(Trim(txtisland.Text)) = LCase("Volcano") Then
        Open App.Path & "\volcano.txt" For Input As #1
        Do Until EOF(1)
            Input #1, Volcano
            lbloutput = lbloutput + Volcano
        Loop
        Close #1
        picisland.Picture = LoadPicture("volcano.jpg")
    ElseIf LCase(Trim(txtisland.Text)) = LCase("Black Morning Spider") Then
        Open App.Path & "\spider.txt" For Input As #1
        Do Until EOF(1)
            Input #1, Spider
            lbloutput = lbloutput + Spider
        Loop
        Close #1
        picisland.Picture = LoadPicture("spider.jpg")
    ElseIf LCase(Trim(txtisland.Text)) = LCase("Totem Pole") Then
        Open App.Path & "\totem.txt" For Input As #1
        Do Until EOF(1)
            Input #1, Totempole
            lbloutput = lbloutput + Totempole
        Loop
        Close #1
        picisland.Picture = LoadPicture("totem.jpg")
    ElseIf LCase(Trim(txtisland.Text)) = LCase("Supply Hut") Then
        Open App.Path & "\hut.txt" For Input As #1
        Do Until EOF(1)
            Input #1, Hut
            lbloutput = lbloutput + Hut
        Loop
        Close #1
        picisland.Picture = LoadPicture("hut.jpg")
    ElseIf LCase(Trim(txtisland.Text)) = LCase("Quick sand") Then
        Open App.Path & "\quicksand.txt" For Input As #1
        Do Until EOF(1)
            Input #1, Quicksand
            lbloutput = lbloutput + Quicksand
        Loop
        Close #1
        picisland.Picture = LoadPicture("quicksand.jpg")
    ElseIf LCase(Trim(txtisland.Text)) = LCase("Bowling Alley") Then
        Open App.Path & "\bowling.txt" For Input As #1
        Do Until EOF(1)
            Input #1, Bowling
            lbloutput = lbloutput + Bowling
        Loop
        Close #1
        picisland.Picture = LoadPicture("bowling.bmp")
    Else
'in case user misspells a word
        MsgBox "That is not a part of the island!", vbOKOnly, "Retype please!"
    End If
End Sub


Private Sub cmdmainpage_Click()
'load main page
    Form1.Show
    Form3.Hide
End Sub
