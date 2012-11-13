VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   8820
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11955
   FillColor       =   &H8000000B&
   LinkTopic       =   "Form1"
   Picture         =   "ScrubsIntro.frx":0000
   ScaleHeight     =   8820
   ScaleWidth      =   11955
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Enter Sacred Heart"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Forte"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   9960
      Picture         =   "ScrubsIntro.frx":11E9E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6120
      Width           =   1815
   End
   Begin VB.CommandButton cmdSignIn 
      BackColor       =   &H00C0C000&
      Caption         =   "Sign In"
      DisabledPicture =   "ScrubsIntro.frx":1328B
      DownPicture     =   "ScrubsIntro.frx":14678
      BeginProperty Font 
         Name            =   "Forte"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      MaskColor       =   &H00FF0000&
      Picture         =   "ScrubsIntro.frx":15A65
      TabIndex        =   0
      Top             =   7560
      UseMaskColor    =   -1  'True
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome to Sacred Heart Hospital!"
      BeginProperty Font 
         Name            =   "Forte"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   11535
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Scrubs Project
'Main/Opening Form (frmMain)
'Ann Boeckmann
'October 25, 2008
'The purpose of this project is to allow Scrubs fans to learn more about the show and partake in
'activities related to the show
'This opening form allows a user to sign in and be welcomed as a doctor


Private Sub cmdEnter_Click()
'allows the user to enter the main menu

frmMain.Hide
frmOptions.Show

End Sub

Private Sub cmdSignIn_Click()
Dim UserName As String
'allows the user to sign in as a doctor

UserName = InputBox("What is your name?", "Name?")
MsgBox "Welcome Dr. " & UserName & "!", , "Welcome"

'The button that opens the main menu is disabled until the user signs in
'Once the user signs in the sign in button disappears and the enter button is enabled

cmdSignIn.Visible = False

cmdEnter.Enabled = True


End Sub


