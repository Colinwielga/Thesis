VERSION 5.00
Begin VB.Form frmWelcome 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Joe's Citation Creator"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8010
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   8010
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtUsersName 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5160
      TabIndex        =   9
      Top             =   2640
      Width           =   2295
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load a Saved Bibliography"
      Height          =   735
      Left            =   5400
      TabIndex        =   6
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   7080
      Picture         =   "frmWelcome.frx":0000
      TabIndex        =   5
      Top             =   5040
      Width           =   855
   End
   Begin VB.CommandButton cmdAPA 
      Caption         =   "APA"
      Height          =   735
      Left            =   3240
      TabIndex        =   2
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton cmdMLA 
      Caption         =   "MLA"
      Height          =   735
      Left            =   840
      TabIndex        =   1
      Top             =   4200
      Width           =   1695
   End
   Begin VB.Label lblEnterName 
      BackColor       =   &H00C0FFC0&
      Caption         =   "First, enter your first and last name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   8
      Top             =   2760
      Width           =   4455
   End
   Begin VB.Label lblCreatedBy 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Created by Joe Sass"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      Top             =   5160
      Width           =   2655
   End
   Begin VB.Label lblExplanation 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "This program allows the user to easily create correctly formatted bibliographies in both MLA and APA format. "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1440
      TabIndex        =   4
      Top             =   960
      Width           =   5295
   End
   Begin VB.Label lblWelcome 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "Welcome to Joe's Citation Creator!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   -120
      TabIndex        =   3
      Top             =   120
      Width           =   8175
   End
   Begin VB.Label lblPleaseChoose 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "Second, please choose the type of bibliography that you wish to create:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      TabIndex        =   0
      Top             =   3480
      Width           =   5535
   End
End
Attribute VB_Name = "frmWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'this is the welcome form, the first form to be displayed

Private Sub cmdAPA_Click()
    'Switches to the APA book input form
    frmWelcome.Hide
    frmAPABook.Show
    'tells sets that it is APA format for formating later on the works cited form
    MLA = False
    ctr = 0
    'sets the user's name to the text in the text box
    UsersName = txtUsersName.Text
End Sub

Private Sub cmdExit_Click()
    'quits the program
    End
End Sub

Private Sub cmdLoad_Click()
    'loads a previously saved array
    ctr = 0
    
    'asks the user to input the name of the file that they have saved previously
    fileName = InputBox("Please type the name of the file that you wish to open (you do not need to type '.txt')", "Open")
    
    'opens the file
    Open App.Path & "\saved\" & fileName & ".txt" For Input As #1
    
    'loads file into an array
    Do Until EOF(1)
        ctr = ctr + 1
        Input #1, AuthorsLastName(ctr), AuthorsFirstName(ctr), AuthorsMiddleName(ctr), Title(ctr), CityPublished(ctr), Publisher(ctr), Year(ctr)
    Loop
    Close #1
    
    'sets the user's name to the text in the text box
    UsersName = txtUsersName.Text
    
    'shows the loaded form, and asks the user for the type of bibliography that they are opening
    frmLoaded.Show
    frmChooseType.Show
    frmWelcome.Hide
End Sub

Private Sub cmdMLA_Click()
    'shows the MLA input form
    frmWelcome.Hide
    frmMLABook.Show
    'sets the type of format to MLA to be used later in the Works cited form
    MLA = True
    ctr = 0
    'defines the user's name as the text in the text box
    UsersName = txtUsersName.Text
End Sub


