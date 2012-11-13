VERSION 5.00
Begin VB.Form frmQuiz 
   BackColor       =   &H000000FF&
   ClientHeight    =   7935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8160
   LinkTopic       =   "Form1"
   ScaleHeight     =   7935
   ScaleWidth      =   8160
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdInfo 
      Caption         =   "More Info"
      Height          =   735
      Left            =   840
      TabIndex        =   13
      Top             =   6840
      Width           =   1695
   End
   Begin VB.PictureBox picOutput 
      Height          =   5415
      Left            =   3840
      ScaleHeight     =   5355
      ScaleWidth      =   3795
      TabIndex        =   12
      Top             =   600
      Width           =   3855
   End
   Begin VB.OptionButton optCA 
      BackColor       =   &H000000FF&
      Caption         =   "Would live in California."
      Height          =   615
      Left            =   600
      TabIndex        =   11
      Top             =   5040
      Width           =   2775
   End
   Begin VB.OptionButton optSaveWorld 
      BackColor       =   &H000000FF&
      Caption         =   "I want to save the world."
      Height          =   735
      Left            =   600
      TabIndex        =   10
      Top             =   4320
      Width           =   3015
   End
   Begin VB.CheckBox chkBigOne 
      BackColor       =   &H000000FF&
      Caption         =   "A Big One"
      Height          =   495
      Left            =   1080
      TabIndex        =   9
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CheckBox chkNoGun 
      BackColor       =   &H000000FF&
      Caption         =   "I don't need one"
      Height          =   495
      Left            =   1080
      TabIndex        =   8
      Top             =   1800
      Width           =   1455
   End
   Begin VB.OptionButton optReadMind 
      BackColor       =   &H000000FF&
      Caption         =   "Sometimes I think I can read minds."
      Height          =   735
      Left            =   600
      TabIndex        =   7
      Top             =   3600
      Width           =   2895
   End
   Begin VB.OptionButton optHummer 
      BackColor       =   &H000000FF&
      Caption         =   "Would love to drive a Hummer!"
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   3000
      Width           =   2775
   End
   Begin VB.CheckBox chkHandgun 
      BackColor       =   &H000000FF&
      Caption         =   "Any Handgun"
      Height          =   495
      Left            =   1080
      TabIndex        =   3
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   735
      Left            =   5520
      TabIndex        =   2
      Top             =   6840
      Width           =   1695
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Patrol"
      Height          =   735
      Left            =   3240
      TabIndex        =   1
      Top             =   6840
      Width           =   1695
   End
   Begin VB.CommandButton cmdResults 
      Caption         =   "Get Results"
      Height          =   735
      Left            =   840
      TabIndex        =   0
      Top             =   5760
      Width           =   1695
   End
   Begin VB.Label lblOption 
      BackColor       =   &H000000FF&
      Caption         =   "Please Select One:"
      Height          =   615
      Left            =   600
      TabIndex        =   6
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Label lblGun 
      BackColor       =   &H000000FF&
      Caption         =   "Please Select Gun of Choice (Select One):"
      Height          =   495
      Left            =   720
      TabIndex        =   5
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "frmQuiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Profile As String
' Patrol Car
' frmQuiz
' Kevin Conlon
' 11/4/08
' interactive game

Private Sub cmdInfo_Click()
    Dim Data As String
    
    Select Case Profile
        Case "James Carter"
            Open App.Path & "\Carter.txt" For Input As #1
            Do Until EOF(1)
                Input #1, Data
            Loop
            Close #1
        Case "Horatio Caine"
            Open App.Path & "\Caine.txt" For Input As #2
            Do Until EOF(2)
                Input #2, Data
            Loop
            Close #2
        Case "Carlton Lassiter"
            Open App.Path & "\Lassiter.txt" For Input As #2
            Do Until EOF(3)
                Input #3, Data
            Loop
            Close #3
        Case "Matt Parkman"
            Open App.Path & "\Parkman.txt" For Input As #4
            Do Until EOF(4)
                Input #4, Data
            Loop
            Close #4
        Case "John McClane"
            Open App.Path & "\McClane.txt" For Input As #5
            Do Until EOF(5)
                Input #5, Data
            Loop
            Close #5
    End Select
    picOutput.Print Data
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdResults_Click()
    Dim Hummer As Boolean
    Dim ReadMind As Boolean
    Dim CA As Boolean
    Dim Handgun As Single
    Dim BigOne As Single
    Dim NoGun As Single
    Dim SaveWorld As Boolean

    Profile = "James Carter"
    Hummer = optHummer.Value
    ReadMind = optReadMind.Value
    CA = optCA.Value
    Handgun = chkHandgun.Value
    BigOne = chkBigOne.Value
    NoGun = chkNoGun.Value
    SaveWorld = optSaveWorld.Value
    
    If Hummer = True And Handgun = 1 Then
        Profile = "Horatio Caine"
    End If
    If ReadMind = True Or SaveWorld = True And Handgun = 1 Then
        Profile = "Matt Parkman"
    End If
    If NoGun = 1 And CA = False Then
        Profile = "John McClane"
        
    End If
    If BigOne = 1 And CA = True Then
        Profile = "Carlton Lassiter"
    End If
    picOutput.Print Profile
End Sub

Private Sub cmdReturn_Click()
    frmPatrolCar.Show
    frmQuiz.Hide
    
End Sub

