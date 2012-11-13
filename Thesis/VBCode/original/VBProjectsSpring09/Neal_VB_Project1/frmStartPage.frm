VERSION 5.00
Begin VB.Form frmStartPage 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   Caption         =   "Start Page"
   ClientHeight    =   7905
   ClientLeft      =   4200
   ClientTop       =   3285
   ClientWidth     =   10905
   DrawMode        =   2  'Blackness
   LinkTopic       =   "Form1"
   ScaleHeight     =   7905
   ScaleWidth      =   10905
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H0080C0FF&
      Caption         =   "Quit"
      Height          =   615
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7080
      Width           =   1935
   End
   Begin VB.CommandButton cmdRecords 
      BackColor       =   &H0080C0FF&
      Caption         =   "SJU Records"
      Height          =   615
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7080
      Width           =   1935
   End
   Begin VB.CommandButton cmdRoster 
      BackColor       =   &H0080C0FF&
      Caption         =   "Roster"
      Height          =   615
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7080
      Width           =   1935
   End
   Begin VB.CommandButton cmdBirkie 
      BackColor       =   &H0080C0FF&
      Caption         =   "Birkebeiner / Kortelopet Results"
      Height          =   615
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7080
      Width           =   1935
   End
   Begin VB.Image imgTeamPic 
      Height          =   4725
      Left            =   1680
      Picture         =   "frmStartPage.frx":0000
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   7500
   End
   Begin VB.Image imgSJU 
      Height          =   2145
      Left            =   240
      Picture         =   "frmStartPage.frx":837A
      Top             =   240
      Width           =   3660
   End
   Begin VB.Label lblNordicSkiing 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Nordic Skiing"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1215
      Left            =   4200
      TabIndex        =   0
      Top             =   600
      Width           =   6135
   End
End
Attribute VB_Name = "frmStartPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project: SJU_Ski_Team
'Form: frmStartPage
'Author: Kevin Neal
'Written: March 20, 2009
'Object:1)Create a project that appeals to my interets and inform others about the team
        '2)Make buttons and a clear interface that allows the user to navigate easily
        '3)Express my knowledge of various VB features
        '4)Make sure other forms are loaded properly when they are shown


Private Sub cmdBirkie_Click()
    'Switch forms to frmBirkie
    frmBirkie.Show
    frmStartPage.Hide
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdRecords_Click()
    'Switches to frmRecords
    frmRecords.Show
    frmStartPage.Hide
End Sub

Private Sub cmdRoster_Click()
    'Switching forms to frmRoster and opening Roster file
    
    frmStartPage.Hide
    frmRoster.Show
    
    'Loading the Names from file
    Open App.Path & "\Names.txt" For Input As #1
    SkierCTR = 0
    Do Until EOF(1)
        SkierCTR = SkierCTR + 1
        Input #1, SkierNames(SkierCTR), SkierGrades(SkierCTR), NumGrade(SkierCTR), SkierScore(SkierCTR)
        Loop
    Close #1
End Sub
