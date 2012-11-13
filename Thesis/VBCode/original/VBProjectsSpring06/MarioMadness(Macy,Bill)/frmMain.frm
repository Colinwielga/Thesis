VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Mario Madness"
   ClientHeight    =   10044
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   10068
   LinkTopic       =   "Form1"
   Picture         =   "frmMain.frx":0000
   ScaleHeight     =   10044
   ScaleWidth      =   10068
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdlogin 
      Caption         =   "Click here to login into the site!"
      Height          =   855
      Left            =   720
      TabIndex        =   8
      Top             =   360
      Width           =   2292
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search Through Information About the Characters and Video Games"
      Height          =   852
      Left            =   720
      TabIndex        =   7
      Top             =   3600
      Width           =   2292
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   615
      Left            =   6720
      TabIndex        =   6
      Top             =   9240
      Width           =   2292
   End
   Begin VB.CommandButton cmdMovingmario 
      Caption         =   "Try your best to hit Mario!"
      Height          =   852
      Left            =   720
      TabIndex        =   5
      Top             =   5760
      Width           =   2292
   End
   Begin VB.CommandButton cmdcharacters 
      Caption         =   "View Brief Profiles on the Characters from the First Game"
      Height          =   852
      Left            =   720
      TabIndex        =   4
      Top             =   2520
      Width           =   2292
   End
   Begin VB.CommandButton cmdOrder 
      Caption         =   "Order Mario Apparel and Accessories"
      Height          =   852
      Left            =   720
      TabIndex        =   3
      Top             =   6840
      Width           =   2292
   End
   Begin VB.CommandButton cmdOrderform 
      Caption         =   "Subscribe to Nintendo Power Magazine"
      Height          =   852
      Left            =   720
      TabIndex        =   2
      Top             =   7920
      Width           =   2292
   End
   Begin VB.CommandButton cmdMatching 
      Caption         =   "Mario Matching Game"
      Height          =   852
      Left            =   720
      TabIndex        =   1
      Top             =   4680
      Width           =   2292
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "Learn a little about the history of Mario and the Maker of this Program"
      Height          =   855
      Left            =   720
      TabIndex        =   0
      Top             =   1440
      Width           =   2292
   End
   Begin VB.PictureBox picBackground 
      Height          =   11055
      Left            =   0
      Picture         =   "frmMain.frx":190F1A
      ScaleHeight     =   11004
      ScaleWidth      =   10404
      TabIndex        =   9
      Top             =   -360
      Width           =   10455
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         Caption         =   "By Bill Macy"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   12
         Top             =   10080
         Width           =   1935
      End
      Begin VB.OLE OLE1 
         AutoActivate    =   3  'Automatic
         BackStyle       =   0  'Transparent
         Class           =   "SoundRec"
         Height          =   375
         Left            =   9240
         OleObjectBlob   =   "frmMain.frx":321E34
         TabIndex        =   11
         Top             =   9840
         Width           =   375
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Mario Madness"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   31.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   5760
         TabIndex        =   10
         Top             =   8160
         Width           =   3375
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: Mario Madness
'Form name: frmMain
'Author: Bill Macy
'Date Written: Tuesday March 14th, 2006
'Objective of project:  The overall purpose of my project is to allow the user to learn a little
            'about Mario and also to have some fun playing a few games and looking into
            'different things Nintendo sells for Mario.  It is to have fun and to enjoy
            'the different options that are allowed.  Click a button and it will bring
            'a new form for you to choose many other selections.  Before you are allowed
            'to access and page, however, you must login and create an account
'Objective of form:  This form allows the user to access the entire project.  There are several
            'buttons for the user to select from in making a choice of what to do.
            'They can choose from logging in, looking at history, characters, stats,
            'can submit a magazine subscription, fill out an order form for Mario Merchandise
            'and play two games: Mario Matching and Mario Catcher.
            


Option Explicit

Private Sub cmdAbout_Click()
    If chosenname = loginname And chosenpassword = loginpassword Then       'Checks to see if the user name and password entered at login are correct
        frmMain.Hide        'Hides the main page
        frmHistory.Show     'Shows the histroy page
    End If
    If myvariable = False Then      'If the user hasnt registered, they wont be allowed to look at anything.  Checks to see if the user is registered
        frmHistory.Hide     'Doesnt allow history page to be seen
        MsgBox "You must login in order to view this page.  Please register or login in now.", , "Login Error!"     'tells the user to log in
        frmMain.Show    'brings back the main page so the user can register
    End If
End Sub

Private Sub cmdexit_Click()
    End     'closes the program
End Sub


Private Sub cmdCharacters_Click()
    If chosenname = loginname And chosenpassword = loginpassword Then       'Checks to see if the user name and password entered at login are correct
        frmMain.Hide        'Hides the main page
        frmCharacters.Show  'Hides the Character page
    End If
    If myvariable = False Then      'If the user hasnt registered, they wont be allowed to look at anything.  Checks to see if the user is registered
        frmCharacters.Hide      'Doesnt allow characters page to be seen
        MsgBox "You must login in order to view this page.  Please register or login in now.", , "Login Error!"     'tells the user to log in
        frmMain.Show    'brings back the main page so the user can register
    End If
End Sub

Private Sub cmdLogin_Click()
    frmMain.Hide        'Hides the main form so the user can log in
    frmLogin.Show       'Shows the login page
End Sub

Private Sub cmdMatching_Click()
    If chosenname = loginname And chosenpassword = loginpassword Then       'Checks to see if the user name and password entered at login are correct
        MsgBox "You must click from the top half of the pictures first in order for the game to work. Then select one from the lower half!  Enjoy!", , "Important"      'Explains to the user an important rule for playing the matching game
        frmMain.Hide        'Hides the main page
        frmMatching.Show        'Hides the Matching page
    End If
    If myvariable = False Then      'If the user hasnt registered, they wont be allowed to look at anything.  Checks to see if the user is registered
        frmMatching.Hide        'Doesnt allow matching page to be seen
        MsgBox "You must login in order to view this page.  Please register or login in now.", , "Login Error!"     'tells the user to log in
        frmMain.Show        'brings back the main page so the user can register
    End If
End Sub

Private Sub cmdMovingmario_Click()
    If chosenname = loginname And chosenpassword = loginpassword Then       'Checks to see if the user name and password entered at login are correct
        frmMain.Hide        'Hides the main page
        frmOptions.Show     'Hides the options page
    End If
    If myvariable = False Then
        frmOptions.Hide     'Doesnt allow options page to be seen
        MsgBox "You must login in order to view this page.  Please register or login in now.", , "Login Error!"     'tells the user to log in
        frmMain.Show        'brings back the main page so the user can register
    End If
End Sub

Private Sub cmdOrder_Click()
    If chosenname = loginname And chosenpassword = loginpassword Then       'Checks to see if the user name and password entered at login are correct
        MsgBox "As a special offer, when you order $45 or more through us you will automatically receive a 15% discount off your bill!", , "Just for you!"      'Informs the user of special savings
        frmMain.Hide        'Hides the main page
        frmOrder.Show       'Shows the Order page
    End If
    If myvariable = False Then
        frmOrder.Hide       'Doesnt allow order page to be seen
        MsgBox "You must login in order to view this page.  Please register or login in now.", , "Login Error!"     'tells the user to log in
        frmMain.Show        'brings back the main page so the user can register
    End If
    
End Sub

Private Sub cmdOrderform_Click()
    If chosenname = loginname And chosenpassword = loginpassword Then       'Checks to see if the user name and password entered at login are correct
        frmMain.Hide        'Hides the main page
        frmMagazine.Show        'Shows the magazine page
    End If
    If myvariable = False Then
        frmMagazine.Hide        'Doesnt allow magazine page to be seen
        MsgBox "You must login in order to view this page.  Please register or login in now.", , "Login Error!"     'tells the user to log in
        frmMain.Show        'brings back the main page so the user can register
    End If
End Sub


Private Sub cmdsearch_Click()
    If chosenname = loginname And chosenpassword = loginpassword Then       'Checks to see if the user name and password entered at login are correct
        frmMain.Hide        'Hides the main page
        frmSearch.Show      'Shows the search page
    End If
    If myvariable = False Then
        frmSearch.Hide      'Hides the search page
        MsgBox "You must login in order to view this page.  Please register or login in now.", , "Login Error!"     'tells the user to log in
        frmMain.Show        'brings back the main page so the user can register
    End If
End Sub
