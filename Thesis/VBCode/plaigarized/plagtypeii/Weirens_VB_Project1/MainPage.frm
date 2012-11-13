VERSION 5.00
Begin VB.Form frmMainPage
   BackColor       =   &H00FF0000&
   Caption         =   "Kayla's Radiology Symptom Checker Home Page"
   ClientHeight    =   10470
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5025
   LinkTopic       =   "Form1"
   ScaleHeight     =   10470
   ScaleWidth      =   5025
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCosts
      BackColor       =   &H0080FFFF&
      Caption         =   "Click Here to Read about the Costs for the Radiologic Procedures that Could or Have Been Recommended to You!"
      BeginProperty Font
         Name            =   "Californian FB"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8520
      Width           =   4575
   End
   Begin VB.CommandButton cmdQuit
      BackColor       =   &H0080FFFF&
      Caption         =   "Quit"
      BeginProperty Font
         Name            =   "Californian FB"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      MaskColor       =   &H0080FFFF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9840
      UseMaskColor    =   -1  'True
      Width           =   4575
   End
   Begin VB.CommandButton cmdEnterChecker
      BackColor       =   &H0080FFFF&
      Caption         =   "Click Here to Begin Entering Your Information and to Figure Out the Right Radiologic Procedure for Your Condition! "
      BeginProperty Font
         Name            =   "Californian FB"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      MaskColor       =   &H0080FFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7080
      UseMaskColor    =   -1  'True
      Width           =   4575
   End
   Begin VB.PictureBox Picture1
      Height          =   4575
      Left            =   240
      Picture         =   "MainPage.frx":0000
      ScaleHeight     =   4515
      ScaleWidth      =   4515
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
   Begin VB.OLE OLE1
      Class           =   "SoundRec"
      Height          =   495
      Left            =   3960
      OleObjectBlob   =   "MainPage.frx":41EF2
      SourceDoc       =   "\\ad\homedir$\Students\K\kaweirens\My Documents\My Music\Heartbeat-SoundBible.com-1259974459.wav"
      TabIndex        =   6
      Top             =   5280
      Width           =   615
   End
   Begin VB.Label Label1
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   $"MainPage.frx":115D0A
      BeginProperty Font
         Name            =   "Californian FB"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   1215
      Left            =   240
      TabIndex        =   2
      Top             =   5760
      Width           =   4575
   End
   Begin VB.Label lblWelcome
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Welcome to Kayla's Radiology Symptom Checker!  "
      BeginProperty Font
         Name            =   "Californian FB"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   855
      Left            =   0
      TabIndex        =   1
      Top             =   4800
      Width           =   4935
   End
End
Attribute VB_Name = "frmMainPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Kayla's Radiology Symptom Checker
'HeadSkull
'Kayla Weirens
'February 15th,2010
'The purpose of this project is to allow for the user to enter their personal information as if they were using a true medical database like WebMD and then enter there symptoms; however, unlike other symptom checkers this recommends a radiologic procedure for them to have done afer visiting with their physician. It is supposed to be fun and informative!
Option Explicit
Private Sub aaaa()
frmMainPage.Hide    'Hides main page
frmCostsPage.Show   'Shows costs page
End Sub
Private Sub cccc()
frmMainPage.Hide    'Hides main page
frmBasicInfo.Show   'Shows basic info page
End Sub
Private Sub ffff()
'Posts message box when quitting program
MsgBox ("Thank You for Using Kayla's Radiology Symptom Checker! I hope that it was able to help and that you feel better soon! :)")
End
End Sub

Private Sub yyyy()
'I got this code from Samantha Arel within her Sample VB right up which I found to be incredibly helpful for my own layout.  So this is courtesy of Stephanie Arel with the idea and code but I changed the numbers for my own preferences.
Top = Screen.Height / 3 - Height / 3
Left = Screen.Width / 3 - Width / 3

End Sub

Private Sub ssss(HEYTHING As Long)
'I got my sound clip from http://soundbible.com/34-Heartbeat.html
End Sub

Private Sub eeee()
'I got this picture from http://www.clker.com/clipart-28058.html
End Sub
