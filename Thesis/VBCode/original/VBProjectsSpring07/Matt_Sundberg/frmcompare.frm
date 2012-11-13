VERSION 5.00
Begin VB.Form frmcompare 
   BackColor       =   &H00008000&
   Caption         =   "Olympic Comparisons"
   ClientHeight    =   5910
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8970
   ForeColor       =   &H00008000&
   LinkTopic       =   "Form1"
   Picture         =   "frmcompare.frx":0000
   ScaleHeight     =   5910
   ScaleWidth      =   8970
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdback 
      BackColor       =   &H00008000&
      Caption         =   "Back To Main Menu"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5760
      Width           =   3135
   End
   Begin VB.PictureBox picResult 
      BackColor       =   &H00008000&
      Height          =   615
      Left            =   1440
      ScaleHeight     =   555
      ScaleWidth      =   8475
      TabIndex        =   1
      Top             =   4920
      Width           =   8535
   End
   Begin VB.CommandButton cmdcomputefastestman 
      Caption         =   "Who's The Fastest Man Of All Time?"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   2880
      Picture         =   "frmcompare.frx":38ADE
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2160
      Width           =   5415
   End
   Begin VB.Label lblClickHere 
      BackColor       =   &H00008000&
      Caption         =   "Click Below For The Answer"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   1560
      Width           =   3975
   End
End
Attribute VB_Name = "frmcompare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    'declare variables
    Dim TimeArray(1 To 80) As Integer
    Dim RunnerArray(1 To 80) As String
    Dim Pass As Integer
    Dim Pos As Integer
    Dim TempTime As Integer
    Dim TempRunner As Integer
    Dim CTR As Integer
    
    'button for navigating back to main menu
Private Sub cmdback_Click()
    frmcompare.Hide
    frmwhichfact.Show
End Sub

Private Sub cmdcomputefastestman_Click()
    
    Pass = 0
    Pos = 0
    CTR = 0
    'read file into two arrays for the runners and their times
    Open App.Path & "\TimeandMan.txt" For Input As #1
    CTR = 0
    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, TimeArray(CTR), RunnerArray(CTR)
    Loop
    Close #1
    'sort by timearray with runnerarray
    For Pass = 1 To CTR
        For Pos = 1 To CTR - 1
            If TimeArray(Pos) < TimeArray(Pos + 1) Then
                TempTime = TimeArray(Pos)
                TimeArray(Pos) = TimeArray(Pos + 1)
                TimeArray(Pos + 1) = TempTime
                TempRunner = RunnerArray(Pos)
                RunnerArray(Pos) = RunnerArray(Pos + 1)
                RunnerArray(Pos + 1) = TempRunner
            End If
        Next Pos
    Next Pass
    
    picResult.Print "The Fastest Runner Ever In The 100 Meter Is", ; RunnerArray(CTR); "Who Ran It In", ; FomatNumber(TimeArray(CTR), 2), "Seconds"
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    , ; "Seconds."
    
End Sub
