VERSION 5.00
Begin VB.Form TorreThree 
   BackColor       =   &H000000FF&
   Caption         =   "500 Time and How Good it is, Getting a computer ID For a Swimming site"
   ClientHeight    =   5205
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8370
   LinkTopic       =   "Form1"
   ScaleHeight     =   5205
   ScaleWidth      =   8370
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdInputName 
      Caption         =   "Input Name"
      Height          =   495
      Left            =   720
      TabIndex        =   8
      Top             =   3960
      Width           =   2175
   End
   Begin VB.CommandButton cmdNextPage 
      Caption         =   "Next Page"
      Height          =   615
      Left            =   5880
      TabIndex        =   6
      Top             =   360
      Width           =   1935
   End
   Begin VB.PictureBox picVanderkaay 
      Height          =   3255
      Left            =   5640
      Picture         =   "frmTorreThree.frx":0000
      ScaleHeight     =   3195
      ScaleWidth      =   2355
      TabIndex        =   2
      Top             =   1800
      Width           =   2415
   End
   Begin VB.CommandButton cmdInput 
      Caption         =   "Input A Time For the 500 Free"
      Height          =   855
      Left            =   2880
      TabIndex        =   1
      Top             =   600
      Width           =   2055
   End
   Begin VB.CommandButton cmdLastPage 
      Caption         =   "Previous Page"
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label lblFindID 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Go To Next Page to Find ID"
      Height          =   255
      Left            =   1080
      TabIndex        =   7
      Top             =   4560
      Width           =   2175
   End
   Begin VB.Label lblEnter 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Enter your first name, middle initial, and last name."
      Height          =   615
      Left            =   360
      TabIndex        =   5
      Top             =   3000
      Width           =   2415
   End
   Begin VB.Label lblID 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Find a user name from your first middle and last name to enter a swimming website."
      Height          =   615
      Left            =   600
      TabIndex        =   4
      Top             =   1920
      Width           =   2895
   End
   Begin VB.Label lblVanderkaay 
      BackColor       =   &H00C0FFC0&
      Caption         =   "500 Free Record Holder Peter Vanderkaay"
      Height          =   375
      Left            =   6000
      TabIndex        =   3
      Top             =   1200
      Width           =   1815
   End
End
Attribute VB_Name = "TorreThree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'The Author of this form is Torre Ahlberg on 11/02/ 2006
'The objective of this form is to have people input times for a 500 yard freestyle
'and then have a program tell them how competitive their time are
'This form also uses multiple forms
'And lastly it has the begining of a program that will allow people
'to input their name and recieve a ID name for access to a swimming site

'This Subroutine uses a case statement to allow the user to enter a time for the 500 Freestyle
'The case program then depending on the time will tell the user how competitive they are
Private Sub cmdInput_Click()
    Dim Time As Single
    Time = InputBox("Enter a 500 freestyle time", , "500 Time")
    
    Select Case Time
    Case Is >= 530
        MsgBox "Try Harder", , "500 Time"
    Case 515 To 529
        MsgBox "Averege", , "500 Time"
    Case 500 To 514
        MsgBox "Competitive", , "500 Time"
    Case 445 To 459
        MsgBox "You are Fast", , "500 Time"
    Case 430 To 444
        MsgBox "Smokeing Fast", , "500 Time"
    Case 415 To 429
        MsgBox "NCAA Division 1 Speed", , "500 Time"
    Case 408 To 414
        MsgBox "NCAA Division 1 Top 8 Material", , "500 Time"
    Case Else
        MsgBox "That Time Is Imposible", , "500 Time"
    End Select
    
        
    
End Sub

'This Subroutine allows the user to input their full name into an Inputbox
'that will then be used on the next form to produce and ID name for access to a swimming site
Private Sub cmdInputName_Click()
    YourName = InputBox("Please Write you full name with middle initial", , "Input Name")
End Sub

'This Subroutine lets the user go from the current form back to the previous form
'using the visible true or false method
Private Sub cmdLastPage_Click()
    TorreThree.Visible = False
    frmTorreTwo.Visible = True
    
End Sub

'This Subroutine lets the user go from the current form to the next form
'usin the visible true or false method
Private Sub cmdNextPage_Click()
    TorreThree.Visible = False
    TorreFour.Visible = True
    
End Sub
