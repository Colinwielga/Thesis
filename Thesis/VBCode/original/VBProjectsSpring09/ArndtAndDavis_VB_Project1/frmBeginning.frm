VERSION 5.00
Begin VB.Form frmBeginning 
   BackColor       =   &H00400000&
   Caption         =   "The Game Of Life"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   FillColor       =   &H00800000&
   ForeColor       =   &H00808080&
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   6135
      Left            =   2280
      Picture         =   "frmBeginning.frx":0000
      ScaleHeight     =   6075
      ScaleWidth      =   10875
      TabIndex        =   4
      Top             =   1920
      Width           =   10935
   End
   Begin VB.CommandButton cmdFinerThingsInLife 
      Caption         =   "Finer Things In Life"
      BeginProperty Font 
         Name            =   "Mathematica5"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   9840
      TabIndex        =   3
      Top             =   8640
      Width           =   2535
   End
   Begin VB.CommandButton cmdJobFuture 
      Caption         =   "Job Future"
      BeginProperty Font 
         Name            =   "Mathematica5"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6480
      TabIndex        =   2
      Top             =   8640
      Width           =   2655
   End
   Begin VB.CommandButton cmdFutureLoveLife 
      Caption         =   "Future Love Life"
      BeginProperty Font 
         Name            =   "Mathematica5"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3240
      TabIndex        =   1
      Top             =   8640
      Width           =   2415
   End
   Begin VB.Label lblThree 
      Alignment       =   2  'Center
      Caption         =   "Step Three:"
      BeginProperty Font 
         Name            =   "Harrington"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10440
      TabIndex        =   7
      Top             =   8400
      Width           =   1095
   End
   Begin VB.Label lblTwo 
      Alignment       =   2  'Center
      Caption         =   "Step Two:"
      BeginProperty Font 
         Name            =   "Harrington"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      TabIndex        =   6
      Top             =   8400
      Width           =   1095
   End
   Begin VB.Label lblOne 
      Alignment       =   2  'Center
      Caption         =   "Step One:"
      BeginProperty Font 
         Name            =   "Harrington"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   8400
      Width           =   1095
   End
   Begin VB.Label lblIntro 
      BackColor       =   &H00400000&
      Caption         =   "Click Below To Begin Your Journey On Discovering Your Future!"
      BeginProperty Font 
         Name            =   "Harrington"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   2280
      TabIndex        =   0
      Top             =   960
      Width           =   12015
   End
End
Attribute VB_Name = "frmBeginning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: The Game Of Life
'Form Name: frmBeginning
'Authors: Pam Arndt and Alisa Davis
'Date Written: 3/4/09
'Objective:This is our starting start up form for the project
'Here the user can select which part of their life they would like to see their future: Future Love Life, Future Jobs, and the "Finer Things in Life" (AKA house, cars, education)
'The overall objective for our project is centered around a fortune teller
'We guide the user along in choosing their love life and/or family, their job preference, education, house, and cars.
Option Explicit
Private Sub Form_Load()

UserFirstName = InputBox("To start, please enter your first name", "First Name")
UserLastName = InputBox("Enter your last name", "Last Name")


End Sub

Private Sub cmdFinerThingsInLife_Click()
frmBeginning.Hide
frmTheFinerThingsInLife.Show
End Sub

Private Sub cmdFutureLoveLife_Click()
'Here we are declaring our variables

Dim Choice As Integer

'Here we are giving the user three options for deciding their ideal love life
'We use an input box to allow them to decide whether they're looking for a husband, wife, or no significant other.

Choice = InputBox("Type a 1 if you're looking for a husband, 2 if you're looking for a wife, or 3 if you prefer to remain single.", "Spouse")

'Here we used an else if then statement to direct the user to the correct form of their dream location.

If Choice = 1 Then
        frmBeginning.Hide
        frmHusband.Show
    ElseIf Choice = 2 Then
        frmBeginning.Hide
        frmWife.Show
    ElseIf Choice = 3 Then
        MsgBox "You're so great, you don't need anyone else! Please choose next category.", , "Single and Loving It"
        Spouse = "None"
        ChildName = "Sir " & Left(UserFirstName, 3) + Mid(UserLastName, 3, 4) & ", your pet cat"
    Else: MsgBox "Sorry, you entered an invalid option plese enter either a 1,2, or 3.", , "Oopsiedaisy!"
End If
End Sub

Private Sub cmdJobFuture_Click()
'Here we are declaring our variables

Dim Choice As Integer

'Here we are giving the user three options for deciding their ideal career
'We use an input box to allow them to decide whether they have or plan to complete college.

Choice = InputBox("Type a 1 if you've completed or planning on completing college, or 2 if you've completed little or no college.", "College Plans")

'Here we used an else if then statement to direct the user to the correct form of their dream location.
    If Choice = 1 Then
        frmBeginning.Hide
        frmCompletedCollege.Show
    ElseIf Choice = 2 Then
        frmBeginning.Hide
        frmNoCollege.Show
    Else: MsgBox "Sorry, you entered an invalid option plese enter either a 1 or 2.", , "Error"
End If
End Sub

