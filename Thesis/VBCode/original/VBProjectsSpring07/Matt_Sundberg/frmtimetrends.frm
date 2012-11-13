VERSION 5.00
Begin VB.Form frmtimetrends 
   BackColor       =   &H00008000&
   Caption         =   "Time Trends Through History"
   ClientHeight    =   6705
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12210
   LinkTopic       =   "Form1"
   Picture         =   "frmtimetrends.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdclear 
      BackColor       =   &H00008000&
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4560
      Width           =   1695
   End
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
      Height          =   1815
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4560
      Width           =   1695
   End
   Begin VB.PictureBox picResult 
      BackColor       =   &H00008000&
      Height          =   5775
      Left            =   4200
      ScaleHeight     =   5715
      ScaleWidth      =   6795
      TabIndex        =   1
      Top             =   600
      Width           =   6855
   End
   Begin VB.CommandButton cmdtrends 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Compute Correlation Between Time And Change In Olympic Runners' 100 Meter Dash Times"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   120
      Picture         =   "frmtimetrends.frx":38ADE
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   600
      Width           =   3735
   End
End
Attribute VB_Name = "frmtimetrends"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    'button for navigating back to main menu
Private Sub cmdback_Click()
    frmtimetrends.Hide
    frmwhichfact.Show
End Sub
    'clears the picture box
Private Sub cmdclear_Click()
    picResult.Cls
End Sub
    
Private Sub cmdtrends_Click()
    'declare arrays
    Dim TimeArray(1 To 11) As Integer
    Dim Difference As Integer
    Dim TotalDifference As Integer
    Dim CTR As Integer
    Dim Pos As Integer
    Dim PercentChange As Integer
    Dim TotalPercentChange As Integer
    'read file into an array
    Open App.Path & "\DatesandTimesbyDecade.txt" For Input As #1
    CTR = 0
    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, TimeArray(CTR)
    Loop
    Close #1
    'calculate change in runner time per 2 or so olympic cycles
    For Pos = 1 To CTR
        Difference = TimeArray(Pos) - TimeArray(Pos + 1)
        picResult.Print "The Difference In Run Time Over Approximately Two Olympic Cycles Is:"; FormatNumber(Difference, 2); "Seconds"
    Next Pos
    
    'calculate total difference over 110 years
    TotalDifference = TimeArray(1) - TimeArray(10)
    'display in a picture box
    picResult.Print "The Total Difference Is:"; FormatNumber(TotalDifference, 2)
    
End Sub
