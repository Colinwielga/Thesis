VERSION 5.00
Begin VB.Form frmOpen 
   BackColor       =   &H00C0C000&
   Caption         =   "Opening"
   ClientHeight    =   8745
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14970
   LinkTopic       =   "Form1"
   ScaleHeight     =   8745
   ScaleWidth      =   14970
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   4695
      Left            =   7440
      Picture         =   "VB project opening.frx":0000
      ScaleHeight     =   4635
      ScaleWidth      =   6915
      TabIndex        =   5
      Top             =   2040
      Width           =   6975
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H0000FFFF&
      Caption         =   "Quit"
      Height          =   615
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7920
      Width           =   1575
   End
   Begin VB.PictureBox Picture2 
      Height          =   4575
      Left            =   120
      Picture         =   "VB project opening.frx":D1EB
      ScaleHeight     =   4515
      ScaleWidth      =   6795
      TabIndex        =   3
      Top             =   3360
      Width           =   6855
   End
   Begin VB.CommandButton cmdNY 
      BackColor       =   &H0000FF00&
      Caption         =   "Click here for New York!"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7080
      Width           =   3135
   End
   Begin VB.CommandButton cmdLA 
      BackColor       =   &H008080FF&
      Caption         =   "Click here for L.A!"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2160
      Width           =   2775
   End
   Begin VB.Label lblOpening 
      BackColor       =   &H00800000&
      Caption         =   "Start your vacation by clicking the destination you wish to travel to below!!"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   9495
   End
End
Attribute VB_Name = "frmOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name:Travel New York And L.A
'Form Name: frmActivities
'Author: Emily Backes
'Date Written: 3-17-09
'Objective: This form is the welcome page for the project it
'allows the user to pick the place they would like to go
'and then a input box pops up and askes the user to input
'a hotel that they would like to stay at
Option Explicit

Private Sub cmdLA_Click()
'declare all the variables
Dim LA As String

 LA = InputBox("Please type in which hotel you would like to stay at, the Marriott or Hilton")
 'if the user choses Marriott then the home page for
 'the Marriott shows up
 If LA = "Marriott" Then
  frmOpen.Hide
  frmMarriott.Show
  'if the user picks the Hilton then the home page for the
  'hilton will show up
ElseIf LA = "Hilton" Then
  frmHilton.Show
  frmOpen.Hide
  'this message appears if the user doesn't pick one of the options
Else: MsgBox (" Hotel not available, please choose from the ones given")

End If




End Sub

Private Sub cmdNY_Click()
Dim NY As String

NY = InputBox("Please type in which hotel you would like to stay at, the Park Central or Warwick")
'if the user enters the park central then the home page comes up
 If NY = "Park Central" Then
    frmParkCentral.Show
    frmOpen.Hide
'if the user enters the warwick then the home page for that shows up
ElseIf NY = "Warwick" Then
    frmWarwick.Show
    frmOpen.Hide
Else: MsgBox ("The hotel you have entered is not available, please pick one from the list")
End If


End Sub

Private Sub cmdquit_Click()
End
End Sub

