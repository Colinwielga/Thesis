VERSION 5.00
Begin VB.Form frmSelectWant 
   BackColor       =   &H0000FFFF&
   Caption         =   "What Are You Looking For?"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9855
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   5865
   ScaleWidth      =   9855
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSubscribe 
      Caption         =   "Subscribe"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   6
      Top             =   4320
      Width           =   2295
   End
   Begin VB.CommandButton cmdAboutMe 
      Caption         =   "About Me"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   5
      Top             =   720
      Width           =   2775
   End
   Begin VB.CommandButton cmdFindGames 
      Caption         =   "Find Games"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      TabIndex        =   3
      Top             =   2520
      Width           =   2775
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   5280
      Width           =   2295
   End
   Begin VB.CommandButton cmdConsoleGuide 
      Caption         =   "Console Guide"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   1
      Top             =   3480
      Width           =   3375
   End
   Begin VB.CommandButton cmdIndustryNews 
      Caption         =   "Industry News"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   0
      Top             =   1560
      Width           =   2775
   End
   Begin VB.Label lblWelcome 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome!!"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      TabIndex        =   4
      Top             =   0
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   4290
      Left            =   840
      Picture         =   "frmSelectWant.frx":0000
      Top             =   240
      Width           =   4290
   End
End
Attribute VB_Name = "frmSelectWant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Chris Orcutt
'frmSelectWant
'26 March 2007
'This form serves as a guide allowing users to move through the program

Option Explicit
Private Sub cmdAboutMe_Click()
    frmSelectWant.Hide      'Hides SelectWant form
    frmAboutMe.Show         'Opens AboutMe form
End Sub
Private Sub cmdConsoleGuide_Click()
    frmSelectWant.Hide      'Hides SelectWant form
    frmConsoleInfo.Show     'Opens ConsoleInfo form
End Sub
Private Sub cmdFindGames_Click()
    frmSelectWant.Hide      'Hides SelectWant form
    frmFindGames.Show       'Opens FindGames form
End Sub
Private Sub cmdGameInfo_Click()
    frmSelectWant.Hide      'Hides SelectWant form
    frmGameInformation.Show 'Opens GameReviews form
End Sub
Private Sub cmdIndustryNews_Click()
    frmSelectWant.Hide      'Hides SelectWant form
    frmIndustryNews.Show    'Opens IndustryNews form
End Sub
Private Sub cmdQuit_Click()
    MsgBox "Thank You For Stopping!!", , "Come Back Soon!"      'Displays message box before user ends session
    End     'Ends program
End Sub
Private Sub cmdSources_Click()
    frmSelectWant.Hide      'Hides SelectWant form
    frmSources.Show         'Opens Sources form
End Sub
Private Sub cmdSubscribe_Click()
    frmSelectWant.Hide      'Hides Selection form
    frmSubscribe.Show       'Shows Subscribe form
End Sub
