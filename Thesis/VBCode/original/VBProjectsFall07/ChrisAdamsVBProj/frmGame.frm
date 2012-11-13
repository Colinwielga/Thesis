VERSION 5.00
Begin VB.Form frmGame 
   Caption         =   "Quest for the Cup"
   ClientHeight    =   8145
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8145
   FillColor       =   &H00008000&
   BeginProperty Font 
      Name            =   "Rockwell Condensed"
      Size            =   18
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   Picture         =   "frmGame.frx":0000
   ScaleHeight     =   8145
   ScaleWidth      =   8145
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picWild 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8175
      Left            =   0
      Picture         =   "frmGame.frx":BC52
      ScaleHeight     =   8115
      ScaleWidth      =   8115
      TabIndex        =   0
      Top             =   0
      Width           =   8175
      Begin VB.PictureBox picAeros 
         Height          =   2535
         Left            =   2640
         Picture         =   "frmGame.frx":16485
         ScaleHeight     =   2475
         ScaleWidth      =   2595
         TabIndex        =   29
         Top             =   0
         Width           =   2655
      End
      Begin VB.CommandButton cmdQuestion 
         Caption         =   "First Question"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   11
         Top             =   7320
         Width           =   1695
      End
      Begin VB.CommandButton cmdQuit 
         Caption         =   "Quit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6240
         TabIndex        =   10
         Top             =   7320
         Width           =   1695
      End
      Begin VB.CommandButton cmdD 
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3240
         MaskColor       =   &H8000000F&
         TabIndex        =   9
         Top             =   6120
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton cmdC 
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   840
         MaskColor       =   &H8000000F&
         TabIndex        =   8
         Top             =   6120
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton cmdB 
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3240
         MaskColor       =   &H8000000F&
         TabIndex        =   7
         Top             =   5040
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton cmdA 
         Appearance      =   0  'Flat
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   840
         MaskColor       =   &H8000000F&
         TabIndex        =   6
         Top             =   5040
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.PictureBox picQuestion 
         BackColor       =   &H00C0FFFF&
         FillColor       =   &H00008000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   120
         ScaleHeight     =   1755
         ScaleWidth      =   5715
         TabIndex        =   5
         Top             =   2760
         Width           =   5775
      End
      Begin VB.PictureBox picConnSmythe 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   240
         Picture         =   "frmGame.frx":18EBE
         ScaleHeight     =   1395
         ScaleWidth      =   1515
         TabIndex        =   4
         Top             =   120
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.PictureBox picStanleyCup 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   2160
         Picture         =   "frmGame.frx":19CA0
         ScaleHeight     =   1395
         ScaleWidth      =   1515
         TabIndex        =   3
         Top             =   120
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.PictureBox picAllStar 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   4200
         Picture         =   "frmGame.frx":1A5B1
         ScaleHeight     =   1395
         ScaleWidth      =   1635
         TabIndex        =   2
         Top             =   120
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.PictureBox picCapt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   6240
         Picture         =   "frmGame.frx":1B935
         ScaleHeight     =   1395
         ScaleWidth      =   1515
         TabIndex        =   1
         Top             =   120
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label lblLad0 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6000
         TabIndex        =   28
         Top             =   5880
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label lblLad1 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6000
         TabIndex        =   27
         Top             =   5640
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label lblLad2 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6000
         TabIndex        =   26
         Top             =   5400
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label lblLad3 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6000
         TabIndex        =   25
         Top             =   5160
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label lblLad4 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6000
         TabIndex        =   24
         Top             =   4920
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label lblLad5 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6000
         TabIndex        =   23
         Top             =   4680
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label lblLad6 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6000
         TabIndex        =   22
         Top             =   4440
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label lblLad7 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6000
         TabIndex        =   21
         Top             =   4200
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label lblLad8 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6000
         TabIndex        =   20
         Top             =   3960
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label lblLad9 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6000
         TabIndex        =   19
         Top             =   3720
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label lblLad10 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6000
         TabIndex        =   18
         Top             =   3480
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label lblLad11 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6000
         TabIndex        =   17
         Top             =   3240
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label lblLad12 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6000
         TabIndex        =   16
         Top             =   3000
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label lblLad13 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6000
         TabIndex        =   15
         Top             =   2760
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label lblLad14 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6000
         TabIndex        =   14
         Top             =   2520
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label lblLad15 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6000
         TabIndex        =   12
         Top             =   2280
         Visible         =   0   'False
         Width           =   1935
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Quest for The Cup~Minnesota Wild Trivia Game

'Author: Chris Adams

'Date: November 2007

'This form is the main game screen

Option Explicit
'Declare all form level variables
Dim CTR As Integer
Dim DProg(1 To 100) As String
Dim FProg(1 To 100) As String
Dim GProg(1 To 100) As String
Dim Spot As Integer
Dim Answer As String
Dim RightAnswer(1 To 15) As String
Dim Question(1 To 15) As String
Dim ACap(1 To 15) As String
Dim BCap(1 To 15)  As String
Dim CCap(1 To 15)  As String
Dim DCap(1 To 15)  As String
Dim QNum As Integer



Private Sub cmdA_Click()    'Sets value for command button "A"

Answer = "A"

    If Answer = RightAnswer(CTR) Then
            Call Right
        Else
            Call Wrong
        End If
        
End Sub
Private Sub cmdB_Click()    'Sets value for command button "B"

Answer = "B"

    If Answer = RightAnswer(CTR) Then
            Call Right
        Else
            Call Wrong
        End If
        
End Sub

Private Sub cmdC_Click()    'Sets value for command button "C"

Answer = "C"

    If Answer = RightAnswer(CTR) Then
            Call Right
        Else
            Call Wrong
        End If

End Sub

Private Sub cmdD_Click()    'Sets value for command button "D"

Answer = "D"

    If Answer = RightAnswer(CTR) Then
            Call Right
        Else
            Call Wrong
        End If
        
End Sub

Private Sub cmdQuestion_Click()

    'Makes the four command buttons appear
    cmdA.Visible = True
    cmdB.Visible = True
    cmdC.Visible = True
    cmdD.Visible = True
    cmdQuestion.Visible = False     'Makes the Question button false

Dim Inc As Integer  'Inc is the incrementer

'open the file containing questions, choices, and correct answer
Open App.Path & "\Files\Questions.txt" For Input As #1
    For Inc = 1 To 15
        Input #1, Question(Inc), ACap(Inc), BCap(Inc), CCap(Inc), DCap(Inc), RightAnswer(Inc)
    Next Inc
    Close #1
    
    'Establish the number of questions based on the position selected by the user
    If Game = 1 Then
       QNum = 15
       Call Ladder1
    ElseIf Game = 2 Then
        QNum = 7
        Call Ladder2
    Else
        QNum = 5
        Call Ladder3
    End If
    
    'Display Questions and captions from the previously opened file
    If CTR = 1 And CTR <= QNum Then
        picQuestion.Print (CTR & "." & Question(CTR))
        cmdA.Caption = ACap(CTR)
        cmdB.Caption = BCap(CTR)
        cmdC.Caption = CCap(CTR)
        cmdD.Caption = DCap(CTR)
    ElseIf CTR = 2 And CTR <= QNum Then
        picQuestion.Print (CTR & "." & Question(CTR))
        cmdA.Caption = ACap(CTR)
        cmdB.Caption = BCap(CTR)
        cmdC.Caption = CCap(CTR)
        cmdD.Caption = DCap(CTR)
    ElseIf CTR = 3 And CTR <= QNum Then
        picQuestion.Print (CTR & "." & Question(CTR))
        cmdA.Caption = ACap(CTR)
        cmdB.Caption = BCap(CTR)
        cmdC.Caption = CCap(CTR)
        cmdD.Caption = DCap(CTR)
    ElseIf CTR = 4 And CTR <= QNum Then
        picQuestion.Print (CTR & "." & Question(CTR))
        cmdA.Caption = ACap(CTR)
        cmdB.Caption = BCap(CTR)
        cmdC.Caption = CCap(CTR)
        cmdD.Caption = DCap(CTR)
    ElseIf CTR = 5 And CTR <= QNum Then
        picQuestion.Print (CTR & "." & Question(CTR))
        cmdA.Caption = ACap(CTR)
        cmdB.Caption = BCap(CTR)
        cmdC.Caption = CCap(CTR)
        cmdD.Caption = DCap(CTR)
    ElseIf CTR = 6 And CTR <= QNum Then
        picQuestion.Print (CTR & "." & Question(CTR))
        cmdA.Caption = ACap(CTR)
        cmdB.Caption = BCap(CTR)
        cmdC.Caption = CCap(CTR)
        cmdD.Caption = DCap(CTR)
    ElseIf CTR = 7 And CTR <= QNum Then
        picQuestion.Print (CTR & "." & Question(CTR))
        cmdA.Caption = ACap(CTR)
        cmdB.Caption = BCap(CTR)
        cmdC.Caption = CCap(CTR)
        cmdD.Caption = DCap(CTR)
    ElseIf CTR = 8 And CTR <= QNum Then
        picQuestion.Print (CTR & "." & Question(CTR))
        cmdA.Caption = ACap(CTR)
        cmdB.Caption = BCap(CTR)
        cmdC.Caption = CCap(CTR)
        cmdD.Caption = DCap(CTR)
    ElseIf CTR = 9 And CTR <= QNum Then
        picQuestion.Print (CTR & "." & Question(CTR))
        cmdA.Caption = ACap(CTR)
        cmdB.Caption = BCap(CTR)
        cmdC.Caption = CCap(CTR)
        cmdD.Caption = DCap(CTR)
    ElseIf CTR = 10 And CTR <= QNum Then
        picQuestion.Print (CTR & "." & Question(CTR))
        cmdA.Caption = ACap(CTR)
        cmdB.Caption = BCap(CTR)
        cmdC.Caption = CCap(CTR)
        cmdD.Caption = DCap(CTR)
    ElseIf CTR = 11 And CTR <= QNum Then
        picQuestion.Print (CTR & "." & Question(CTR))
        cmdA.Caption = ACap(CTR)
        cmdB.Caption = BCap(CTR)
        cmdC.Caption = CCap(CTR)
        cmdD.Caption = DCap(CTR)
    ElseIf CTR = 12 And CTR <= QNum Then
        picQuestion.Print (CTR & "." & Question(CTR))
        cmdA.Caption = ACap(CTR)
        cmdB.Caption = BCap(CTR)
        cmdC.Caption = CCap(CTR)
        cmdD.Caption = DCap(CTR)
    ElseIf CTR = 13 And CTR <= QNum Then
        picQuestion.Print (CTR & "." & Question(CTR))
        cmdA.Caption = ACap(CTR)
        cmdB.Caption = BCap(CTR)
        cmdC.Caption = CCap(CTR)
        cmdD.Caption = DCap(CTR)
    ElseIf CTR = 14 And CTR <= QNum Then
        picQuestion.Print (CTR & "." & Question(CTR))
        cmdA.Caption = ACap(CTR)
        cmdB.Caption = BCap(CTR)
        cmdC.Caption = CCap(CTR)
        cmdD.Caption = DCap(CTR)
    ElseIf CTR = 15 And CTR <= QNum Then
        picQuestion.Print (CTR & "." & Question(CTR))
        cmdA.Caption = ACap(CTR)
        cmdB.Caption = BCap(CTR)
        cmdC.Caption = CCap(CTR)
        cmdD.Caption = DCap(CTR)
    End If
    

End Sub

Private Sub Ladder1()   'This establishes the side bar if the user selects the forward position

        lblLad0.Visible = True
        lblLad0.Caption = FProg(16)
        lblLad1.Visible = True
        lblLad1.Caption = FProg(15)
        lblLad2.Visible = True
        lblLad2.Caption = FProg(14)
        lblLad3.Visible = True
        lblLad3.Caption = FProg(13)
        lblLad4.Visible = True
        lblLad4.Caption = FProg(12)
        lblLad5.Visible = True
        lblLad5.Caption = FProg(11)
        lblLad6.Visible = True
        lblLad6.Caption = FProg(10)
        lblLad7.Visible = True
        lblLad7.Caption = FProg(9)
        lblLad8.Visible = True
        lblLad8.Caption = FProg(8)
        lblLad9.Visible = True
        lblLad9.Caption = FProg(7)
        lblLad10.Visible = True
        lblLad10.Caption = FProg(6)
        lblLad11.Visible = True
        lblLad11.Caption = FProg(5)
        lblLad12.Visible = True
        lblLad12.Caption = FProg(4)
        lblLad13.Visible = True
        lblLad13.Caption = FProg(3)
        lblLad14.Visible = True
        lblLad14.Caption = FProg(2)
        lblLad15.Visible = True
        lblLad15.Caption = FProg(1)
        
        If CTR = 0 Then
            lblLad0.BackColor = vbRed
            lblLad0.ForeColor = vbGreen
        ElseIf CTR = 1 Then
            lblLad0.BackColor = vbRed
            lblLad0.ForeColor = vbGreen
        ElseIf CTR = 2 Then
            lblLad1.BackColor = vbRed
            lblLad1.ForeColor = vbGreen
        ElseIf CTR = 3 Then
            lblLad2.BackColor = vbRed
            lblLad2.ForeColor = vbGreen
        ElseIf CTR = 4 Then
            lblLad3.BackColor = vbRed
            lblLad3.ForeColor = vbGreen
        ElseIf CTR = 5 Then
            lblLad4.BackColor = vbRed
            lblLad4.ForeColor = vbGreen
            picAeros.Visible = False
        ElseIf CTR = 6 Then
            lblLad5.BackColor = vbRed
            lblLad5.ForeColor = vbGreen
        ElseIf CTR = 7 Then
            lblLad6.BackColor = vbRed
            lblLad6.ForeColor = vbGreen
        ElseIf CTR = 8 Then
            lblLad7.BackColor = vbRed
            lblLad7.ForeColor = vbGreen
        ElseIf CTR = 9 Then
            lblLad8.BackColor = vbRed
            lblLad8.ForeColor = vbGreen
        ElseIf CTR = 10 Then
            lblLad9.BackColor = vbRed
            lblLad9.ForeColor = vbGreen
            picCapt.Visible = True
        ElseIf CTR = 11 Then
            lblLad10.BackColor = vbRed
            lblLad10.ForeColor = vbGreen
            picAllStar.Visible = True
        ElseIf CTR = 12 Then
            lblLad11.BackColor = vbRed
            lblLad11.ForeColor = vbGreen
        ElseIf CTR = 13 Then
            lblLad12.BackColor = vbRed
            lblLad12.ForeColor = vbGreen
        ElseIf CTR = 14 Then
            lblLad13.BackColor = vbRed
            lblLad13.ForeColor = vbGreen
            picStanleyCup.Visible = True
        ElseIf CTR = 15 Then
            lblLad14.BackColor = vbRed
            lblLad14.ForeColor = vbGreen
            picConnSmythe.Visible = True
        ElseIf CTR = 16 Then
            lblLad15.BackColor = vbRed
            lblLad15.ForeColor = vbGreen
       End If
               
End Sub

Private Sub Ladder2()   'This establishes the side bar if the user selects the defense position

        lblLad0.Visible = True
        lblLad0.Caption = DProg(8)
        lblLad1.Visible = True
        lblLad1.Caption = DProg(7)
        lblLad2.Visible = True
        lblLad2.Caption = DProg(6)
        lblLad3.Visible = True
        lblLad3.Caption = DProg(5)
        lblLad4.Visible = True
        lblLad4.Caption = DProg(4)
        lblLad5.Visible = True
        lblLad5.Caption = DProg(3)
        lblLad6.Visible = True
        lblLad6.Caption = DProg(2)
        lblLad7.Visible = True
        lblLad7.Caption = DProg(1)
        
        
        If CTR = 0 Then
            lblLad0.BackColor = vbRed
            lblLad0.ForeColor = vbGreen
        ElseIf CTR = 1 Then
            lblLad0.BackColor = vbRed
            lblLad0.ForeColor = vbGreen
        ElseIf CTR = 2 Then
            lblLad1.BackColor = vbRed
            lblLad1.ForeColor = vbGreen
        ElseIf CTR = 3 Then
            lblLad2.BackColor = vbRed
            lblLad2.ForeColor = vbGreen
            picAeros.Visible = False
        ElseIf CTR = 4 Then
            lblLad3.BackColor = vbRed
            lblLad3.ForeColor = vbGreen
        ElseIf CTR = 5 Then
            lblLad4.BackColor = vbRed
            lblLad4.ForeColor = vbGreen
        ElseIf CTR = 6 Then
            lblLad5.BackColor = vbRed
            lblLad5.ForeColor = vbGreen
            picAllStar.Visible = True
        ElseIf CTR = 7 Then
            lblLad6.BackColor = vbRed
            lblLad6.ForeColor = vbGreen
            picStanleyCup.Visible = True
        ElseIf CTR = 8 Then
            lblLad7.BackColor = vbRed
            lblLad7.ForeColor = vbGreen
            picConnSmythe.Visible = True
        End If
 
End Sub

Private Sub Ladder3()   'This establishes the side bar if the user selects the goalie position

        lblLad0.Visible = True
        lblLad0.Caption = GProg(6)
        lblLad1.Visible = True
        lblLad1.Caption = GProg(5)
        lblLad2.Visible = True
        lblLad2.Caption = GProg(4)
        lblLad3.Visible = True
        lblLad3.Caption = GProg(3)
        lblLad4.Visible = True
        lblLad4.Caption = GProg(2)
        lblLad5.Visible = True
        lblLad5.Caption = GProg(1)
        
        If CTR = 0 Then
            lblLad0.BackColor = vbRed
            lblLad0.ForeColor = vbGreen
        ElseIf CTR = 1 Then
            lblLad0.BackColor = vbRed
            lblLad0.ForeColor = vbGreen
        ElseIf CTR = 2 Then
            lblLad1.BackColor = vbRed
            lblLad1.ForeColor = vbGreen
            picAeros.Visible = False
        ElseIf CTR = 3 Then
            lblLad2.BackColor = vbRed
            lblLad2.ForeColor = vbGreen
        ElseIf CTR = 4 Then
            lblLad3.BackColor = vbRed
            lblLad3.ForeColor = vbGreen
            picAllStar.Visible = True
        ElseIf CTR = 5 Then
            lblLad4.BackColor = vbRed
            lblLad4.ForeColor = vbGreen
             picStanleyCup.Visible = True
        ElseIf CTR = 6 Then
            lblLad5.BackColor = vbRed
            lblLad5.ForeColor = vbGreen
            picConnSmythe.Visible = True
        End If
        
End Sub
Private Sub cmdQuit_Click()
    
    'Shows form Sources
    frmGame.Hide
    frmSources.Show

End Sub

Private Sub Right()         'This sub form tells the computer what to do in the instance of a right answer
    
    MsgBox "That is the Correct!"
        cmdQuestion.Caption = "Next Question"
        cmdQuestion.Visible = True
        CTR = CTR + 1
        picQuestion.Cls
    If CTR > QNum And QNum > 7 Then
        frmHoF.Show
        frmGame.Hide
    ElseIf CTR > QNum And QNum = 5 Then
        MsgBox "Congratulations " & PlayerFirst & PlayerLast & "! You have won the Conn Smythe Trophy for MVP and are a true member of the team of 18,000!", , "You beat the Easy Level!"
    ElseIf CTR > QNum And QNum = 7 Then
        MsgBox "Congratulations " & PlayerFirst & PlayerLast & "! You have won the Conn Smythe Trophy for MVP and are a true member of the team of 18,000!", , "You beat the Medium Level!"
    End If
        
End Sub
Private Sub Wrong()     'This sub form tells the computer what to do in the instance of a wrong answer
    
    frmBoogey.Show
    frmGame.Hide

End Sub

Private Sub Form_Load()

CTR = 1 'This is the question that the user is on
    'Load the proper values into the side bar from the file based on which position the user selected
    If Game = 1 Then
        Open App.Path & "\Files\FProg.txt" For Input As #2
            Do Until EOF(2)
                Spot = Spot + 1
                Input #2, FProg(Spot)
            Loop
        Close #2
    ElseIf Game = 2 Then
        Open App.Path & "\Files\DProg.txt" For Input As #3
            Do Until EOF(3)
                Spot = Spot + 1
                Input #3, DProg(Spot)
            Loop
        Close #3
    Else
        Open App.Path & "\Files\GProg.txt" For Input As #4
            Do Until EOF(4)
                Spot = Spot + 1
                Input #4, GProg(Spot)
            Loop
        Close #4
    End If
    
End Sub
