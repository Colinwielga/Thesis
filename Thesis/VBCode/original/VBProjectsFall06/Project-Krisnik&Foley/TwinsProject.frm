VERSION 5.00
Begin VB.Form HomePage 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Homepage"
   ClientHeight    =   7800
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10065
   LinkTopic       =   "Form1"
   ScaleHeight     =   7800
   ScaleWidth      =   10065
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H0080C0FF&
      Caption         =   "Quit Program"
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6840
      Width           =   2295
   End
   Begin VB.CommandButton cmdGame 
      BackColor       =   &H0080C0FF&
      Caption         =   "Know Your Twins"
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5160
      Width           =   2295
   End
   Begin VB.CommandButton cmdRoster 
      BackColor       =   &H0080C0FF&
      Caption         =   "Twins Active Roster && Player Statistics"
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2640
      Width           =   2295
   End
   Begin VB.CommandButton cmdManager 
      BackColor       =   &H0080C0FF&
      Caption         =   "Twins Managers"
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3480
      Width           =   2295
   End
   Begin VB.CommandButton cmdApparel 
      BackColor       =   &H0080C0FF&
      Caption         =   "Shop For Twins Apparel"
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6000
      Width           =   2295
   End
   Begin VB.CommandButton cmdDome 
      BackColor       =   &H0080C0FF&
      Caption         =   "Proposed Twins Ballpark"
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4320
      Width           =   2295
   End
   Begin VB.PictureBox Picture2 
      Height          =   2415
      Left            =   120
      Picture         =   "TwinsProject.frx":0000
      ScaleHeight     =   2355
      ScaleWidth      =   2355
      TabIndex        =   1
      Top             =   0
      Width           =   2415
   End
   Begin VB.PictureBox Picture1 
      Height          =   975
      Left            =   4080
      Picture         =   "TwinsProject.frx":EA50
      ScaleHeight     =   915
      ScaleWidth      =   4515
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
   Begin VB.PictureBox Picture3 
      Height          =   7695
      Left            =   2640
      Picture         =   "TwinsProject.frx":103EB
      ScaleHeight     =   7635
      ScaleWidth      =   7395
      TabIndex        =   3
      Top             =   0
      Width           =   7455
      Begin VB.PictureBox picResults 
         BeginProperty Font 
            Name            =   "Eras Bold ITC"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4320
         ScaleHeight     =   435
         ScaleWidth      =   2715
         TabIndex        =   11
         Top             =   6960
         Width           =   2775
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Caption         =   "Click Me!"
         BeginProperty Font 
            Name            =   "Eras Bold ITC"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   6960
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Eras Bold ITC"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   615
         Left            =   240
         TabIndex        =   4
         Text            =   "Created By: Mike Foley & Jake Krisnik"
         Top             =   6480
         Width           =   6975
      End
   End
End
Attribute VB_Name = "HomePage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Project Name: Twins Baseball
' Form Name: Homepage
' Authors: Jake Krisnik & Mike Foley
' Date Written: October 19, 2006
' Overall Objective: For this project we decided to create a program that would provide
'                    the user information and a general overview about the Twins Baseball
'                    Organization. We felt that it would be beneficial to the user to gain
'                    insights on individual players, the Twins Organization, and give them
'                    the capability to browse common Twins merchandise.
' Form Objective: To provide the user with the ability to navigate away from the homepage
'                 to a variety of choices including forms that encompass: the roster, Twins
'                 managers, information on the proposed Twins ballpark, information on Twins
'                 players linked to pictures, shop for Twins apparel, and quit the program.
Option Explicit

Private Sub cmdApparel_Click()
' This command allows the user to navigate to the Twins apparel form while hiding all other
' open forms.
    HomePage.Hide
    TwinsApparel.Show
    Metrodome.Hide
    Managers.Hide
    TeamRoster.Hide
    KnowYourTwins.Hide
End Sub

Private Sub cmdDome_Click()
' This command allows the user to navigate to the proposed Twins Ballpark form while hiding
' all other open forms.
    HomePage.Hide
    Metrodome.Show
    Managers.Hide
    TwinsApparel.Hide
    TeamRoster.Hide
    KnowYourTwins.Hide
End Sub

Private Sub cmdGame_Click()
' This command allows the user to navigate to the know your Twins form while hiding
' all other open forms.
    KnowYourTwins.Show
    HomePage.Hide
    Managers.Hide
    Metrodome.Hide
    TwinsApparel.Hide
    TeamRoster.Hide
End Sub

Private Sub cmdManager_Click()
' This command allows the user to navigate to the Twins managers form while hiding
' all other open forms.
    HomePage.Hide
    Managers.Show
    Metrodome.Hide
    TwinsApparel.Hide
    TeamRoster.Hide
    KnowYourTwins.Hide
End Sub

Private Sub cmdQuit_Click()
' This command allows the user to navigate away from the entire program itself.
    End
End Sub

Private Sub cmdRoster_Click()
' This command allows the user to navigate to the Twins roster form while hiding
' all other open forms.
    HomePage.Hide
    TeamRoster.Show
    Metrodome.Hide
    TwinsApparel.Hide
    Managers.Hide
    KnowYourTwins.Hide
End Sub

Private Sub Command1_Click()
' This code was created for a fun little interaction with the user on the homepage.
' The user is asked to give a number between 1 and 10 to see if they have a shot of
' making the Minnesota Twins team. There are some different promts based on case
' statements which give the user an answer.
    Dim Number As Single, Result As String
    'clear the pictureBox
    picResults.Cls
    MsgBox "Do you think you have a shot at becoming a Minnesota Twins Baseball player?!"
    'This line of code uses an InputBox to get input from the user.
    'The box will pop up and prompt the user for a number.
    Number = InputBox("Enter a number, (0-10) ")
    Select Case Number
        Case Is >= 8
            Result = "For sure"
        Case Is >= 6
            Result = "Definitely"
        Case 4 To 5
            Result = "Not today"                'The select case allows the program to search the array
        Case 3                                  'once it matches the information inputed from the user it
            Result = "Yes"                      'prints out the correct information stored.
        Case Else
            Result = "My sources say maybe"
    End Select
    picResults.Print Result; "."
End Sub

