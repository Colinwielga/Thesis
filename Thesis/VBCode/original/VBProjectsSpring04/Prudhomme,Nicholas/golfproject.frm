VERSION 5.00
Begin VB.Form golfcourse 
   BackColor       =   &H00C00000&
   Caption         =   "Form2"
   ClientHeight    =   5940
   ClientLeft      =   2100
   ClientTop       =   195
   ClientWidth     =   6540
   LinkTopic       =   "Form2"
   ScaleHeight     =   5940
   ScaleWidth      =   6540
   Begin VB.PictureBox Picture1 
      Height          =   4215
      Left            =   4680
      Picture         =   "golf project.frx":0000
      ScaleHeight     =   4155
      ScaleWidth      =   1155
      TabIndex        =   7
      Top             =   240
      Width           =   1215
      Begin VB.Label Label5 
         Caption         =   "4"
         Height          =   255
         Left            =   840
         TabIndex        =   11
         Top             =   1920
         Width           =   135
      End
      Begin VB.Label Label4 
         Caption         =   "3"
         Height          =   255
         Left            =   720
         TabIndex        =   10
         Top             =   840
         Width           =   135
      End
      Begin VB.Label Label3 
         Caption         =   "2"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   960
         Width           =   135
      End
      Begin VB.Label Label2 
         Caption         =   "1"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   1800
         Width           =   135
      End
   End
   Begin VB.CommandButton cmdyard 
      Caption         =   "Yardage"
      Enabled         =   0   'False
      Height          =   615
      Left            =   2160
      TabIndex        =   6
      Top             =   3840
      Width           =   1575
   End
   Begin VB.TextBox txtlocation 
      Height          =   495
      Left            =   2400
      TabIndex        =   4
      Top             =   2880
      Width           =   495
   End
   Begin VB.CommandButton cmdclubs 
      Caption         =   "Back to Clubs!"
      Height          =   615
      Left            =   2160
      TabIndex        =   3
      Top             =   4800
      Width           =   1575
   End
   Begin VB.CommandButton cmdquit2 
      Caption         =   "Quit"
      Height          =   615
      Left            =   4320
      TabIndex        =   2
      Top             =   4800
      Width           =   1575
   End
   Begin VB.CommandButton cmdhdcp 
      BackColor       =   &H00FFFFFF&
      Caption         =   "What's my handicap?"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   4320
      Width           =   1575
   End
   Begin VB.PictureBox pichole 
      BackColor       =   &H00FFFFFF&
      Height          =   2295
      Left            =   240
      ScaleHeight     =   2235
      ScaleWidth      =   3435
      TabIndex        =   0
      Top             =   240
      Width           =   3495
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C00000&
      Caption         =   "By:  Nicholas Prudhomme"
      Height          =   255
      Left            =   2280
      TabIndex        =   12
      Top             =   5640
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Enter a location number and click Yardage."
      Height          =   735
      Left            =   480
      TabIndex        =   5
      Top             =   2880
      Width           =   1335
   End
End
Attribute VB_Name = "golfcourse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Golf_Project (N_Prudhomme_project.vbp)
'File Name: golfcourse "golf project.frm"
'Author: Nicholas Prudhomme
'Date Written: March 12, 2004
'Form Purpose: To teach the user more about location on a golf course
              'Also to teach the user about the handicap system
              'To give advice to fellow golfers
'Option Explicit is a command force that makes the user declare all
'variables before they can be used
Option Explicit
Dim path As String

Private Sub cmdclubs_Click()
'this button will show the clubs section of the project
golfcourse.Hide
golfclub.Show
End Sub

Private Sub cmdhdcp_Click()
'the intention of this button is to give the user the opprotunity to roughly find out what their handicap may be
Dim score As Integer
Dim handicap As Single
Dim ctr As Integer
Dim total As Integer
ctr = 0
'total is used so that we can have an average score, not just one
total = 0
'this is seen in almost all command buttons to clear the result box
pichole.Cls
Do Until ctr = 5
    score = InputBox("Enter a score on 18 holes. (You will do this five times)", "Handicap")
    'score is the number that will partially determine your handicap
    ctr = ctr + 1 'this will make sure the user enters five numbers
    total = score + total
Loop
'print statements to give background on the idea of a handicap
pichole.Print "A handicap is the number of strokes you"
pichole.Print "would shoot above par for a single round."
'the 72.2 and 139 are the slope and rating of Pebble Beach, the number one golf course in the world
handicap = ((total / 5) - 72.2) * 113 / 139
pichole.Print "Your handicap is"; " "; (FormatNumber(Abs(handicap), 1)); "."
pichole.Print
'this section of code will give advice to the user on what to do about their current golf game
pichole.Print "Our advice is that:"
If handicap < 10 Then 'If else statement used to print corresponding advice
        pichole.Print "You are a great golfer."
    ElseIf handicap < 20 Then
        pichole.Print "You should keep working at it, you'll get there."
    ElseIf handicap > 20 Then
        pichole.Print "You should seek a golf professional to help improve your game."
End If
End Sub
'will end the program
Private Sub cmdquit2_Click()
    End
End Sub

Private Sub cmdyard_Click()
'this button will tell you how far you hit your tee shot when you select a specific location
Dim shot As Integer
Dim yardage(1 To 4) As Integer
Dim choice As Integer
'this is number that the user will enter for their location
shot = txtlocation.Text
pichole.Cls
Open path & "yards.txt" For Input As #1
'will load the file into an array
For choice = 1 To 4
    Input #1, yardage(choice)
Next choice
If shot > 4 Then
        'this message box will ensure that the user enters a "correct number"
        MsgBox "Sorry, you must enter a number between 1 and 4.", , "Error"
    ElseIf shot < 1 Then
        MsgBox "Sorry, you must enter a number between 1 and 4.", , "Error"
End If
Select Case shot 'case statement used to match output with user input
    Case Is = 1
        pichole.Print "You hit your drive"; yardage(1); "yards"
    Case 2
        pichole.Print "You hit it"; yardage(2); "yards"; Tab(1); "but you are in the bunker."
    Case 3
        pichole.Print "You hit it"; yardage(3); "yards"; Tab(1); "but you are in the water."
    Case 4
        pichole.Print "You hit it in the rough"; yardage(4); "yards,not bad."
End Select
Close #1 'will close the file
End Sub


Private Sub Form_Load()
    'having the path statement will allow others to access the files on a restricted drive
    path = "N:\CS130\handin\Prudhomme, Nicholas\"
End Sub

Private Sub txtlocation_Change()
'this piece of the code ensures that the user does not try and click the yardage button before they enter a location
If txtlocation <> "" Then
    cmdyard.Enabled = True
    Else: cmdyard.Enabled = False
End If
End Sub
