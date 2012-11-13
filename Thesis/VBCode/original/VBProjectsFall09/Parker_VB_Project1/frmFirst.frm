VERSION 5.00
Begin VB.Form frmFirst 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Begin"
   ClientHeight    =   4680
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9435
   LinkTopic       =   "Form1"
   Picture         =   "frmFirst.frx":0000
   ScaleHeight     =   4680
   ScaleWidth      =   9435
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdResearch 
      Height          =   735
      Left            =   360
      Picture         =   "frmFirst.frx":27D7
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2760
      Width           =   2055
   End
   Begin VB.CommandButton cmdQuit 
      Height          =   735
      Left            =   6600
      Picture         =   "frmFirst.frx":2FCB
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2760
      Width           =   2175
   End
   Begin VB.CommandButton cmdEnter 
      Height          =   735
      Left            =   3480
      Picture         =   "frmFirst.frx":7F35
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2760
      Width           =   2295
   End
   Begin VB.Label lblExit 
      Alignment       =   2  'Center
      Caption         =   "Exit Program"
      Height          =   255
      Left            =   6840
      TabIndex        =   5
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label lblBuild 
      Alignment       =   2  'Center
      Caption         =   "Build your Own Car!"
      Height          =   255
      Left            =   3720
      TabIndex        =   4
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label lblResearch 
      Alignment       =   2  'Center
      Caption         =   "A Little Research"
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   2400
      Width           =   1575
   End
End
Attribute VB_Name = "frmFirst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project Name is WickedFunCarBuilder
'Form Name is frmFirst
'Author is Dan Parker
'Date written 10/12/09
'The purpose of this project to to aide a user in the process of researching/buying
'a new car. The program provides the user with helpful facts about cars in the
'research section, such as a list of the average cost of the cars sold by particular
'automakers. The program also provides the user with a build tool, with which they can
'choose an automotive body style and add options to it. The program will give the user
'the total cost for that car.
'This form is the first form, so it serves as a base for the whole program. From here,
'the user can research or build cars.
Private Sub cmdEnter_Click()
    
    'load the options into an array
    Open App.Path & "\options.txt" For Input As #1
    ctr = 0
    Do Until EOF(1)
        ctr = ctr + 1
        Input #1, priceOption(ctr)
    Loop
    
    Close #1
    
    'load the car prices into an array
    Open App.Path & "\carPrices.txt" For Input As #2
    ctr = 0
    Do While Not EOF(2)
        ctr = ctr + 1
        Input #2, price(ctr)
    Loop
    Close #2 'close file
    
    MsgBox ("Before we build, let's take a look at the different styles that are available to you.")
    
    frmFirst.Hide 'hides Welcome page from user
    frmPictures.Show 'shows next page to user
    
End Sub

Private Sub cmdQuit_Click()
MsgBox ("Thanks for using the Wicked Fun Car Builder, " & " " & UserName & "!")
End 'ends program
End Sub

Private Sub cmdResearch_Click()
    frmFirst.Hide 'hides welcome page form user
    frmResearch.Show 'shows Research page to user
End Sub
Private Sub Form_Load()
'provide an intro to the user and gather the user's name
'the program will not load unless the user enters his or her name into the input box
frmFirst.Hide
Dim found As Boolean
found = False

    
    
Do While found = False
    UserName = InputBox("Welcome to Dan's auto builder. This program will help you build a car and give you an estimate of what that car will cost. Let's begin. What's your name?")
    If Len(UserName) = 0 Then  'this keeps the program from loading if nothing is entered
        MsgBox "Try again.", , "We need your name to continue!"
    Else
        'the user entered at least on character, so a message pops up and the first form is shown
        MsgBox ("Welcome," & " " & UserName & "! Let's begin. To do so, please click the Research button to learn a little about automakers in America, or click the Build button to create your own car.")
        found = True
        frmFirst.Show 'show First page
    End If
Loop
End Sub

