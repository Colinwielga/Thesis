VERSION 5.00
Begin VB.Form frmBeginning 
   BackColor       =   &H00000000&
   Caption         =   "Beginning of Project"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11040
   BeginProperty Font 
      Name            =   "Century Gothic"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "frmBeginning.frx":0000
   ScaleHeight     =   9180
   ScaleWidth      =   11040
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit This Rad Program"
      Height          =   735
      Left            =   9840
      TabIndex        =   6
      Top             =   8280
      Width           =   1095
   End
   Begin VB.CommandButton cmdWorksCited 
      Caption         =   "Works Cited"
      Height          =   495
      Left            =   9000
      TabIndex        =   5
      Top             =   8400
      Width           =   735
   End
   Begin VB.CommandButton cmdWebsite 
      BackColor       =   &H80000007&
      Caption         =   "Official website"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8760
      TabIndex        =   4
      Top             =   7560
      Width           =   2175
   End
   Begin VB.CommandButton cmdTour 
      BackColor       =   &H80000000&
      Caption         =   "Weezer's Troublemaker Tour Schedule"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8760
      TabIndex        =   3
      Top             =   6480
      Width           =   2175
   End
   Begin VB.CommandButton cmdSong 
      Caption         =   "See Weezer's Discography and song information"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8760
      TabIndex        =   2
      Top             =   5400
      Width           =   2175
   End
   Begin VB.CommandButton cmdInfo 
      Caption         =   "Learn about Weezer's band members"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8760
      TabIndex        =   1
      Top             =   4320
      Width           =   2175
   End
   Begin VB.OLE OLE1 
      Class           =   "SoundRec"
      Height          =   495
      Left            =   7920
      OleObjectBlob   =   "frmBeginning.frx":AFCCE
      TabIndex        =   7
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblWeezer 
      Caption         =   $"frmBeginning.frx":1F6EE6
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   8760
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmBeginning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Weezer
'Form Name: frmBeginning.frm
'Author: Emily Balamut
'Date Written: October 24th, 2008
'Objective: The overall ocjective of this project is to act like an official website
'for the alternative rock band, Weezer. I included an information page about the
'members, a tour schedule of their most recent tour, a discography page, a picture
'and concert page, and a song page in the project.
'This form is the first form in the project (besides the introductory form).
'It has a button to take the user to the tour schedule, a button to learn about
'the band members, a button to bring the user to discography information, and a
'button that sends the user to Weezer's official website.
Option Explicit
Private Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub cmdExit_Click()
MsgBox "Thanks for rocking out with Weezer, " & UserName & "! See you later!", , "Bye!"
End
End Sub

Private Sub cmdInfo_Click()
    frmSTART.Hide
    frmInfo.Show
End Sub

Private Sub cmdSong_Click()
    frmSTART.Hide
    frmDisco.Show
End Sub

Private Sub cmdTour_Click()
    frmSTART.Hide
    frmSchedule.Show
End Sub

Private Sub cmdWebsite_Click()
    
    ShellExecute Me.hWnd, "open", "http://www.weezer.com/", "", "", False
    
End Sub
Private Sub cmdWelcome_Click()
    frmSTART.Hide
    frmSchedule.Show
End Sub

Private Sub cmdWorksCited_Click()
    frmSTART.Hide
    frmWorksCited.Show
End Sub

