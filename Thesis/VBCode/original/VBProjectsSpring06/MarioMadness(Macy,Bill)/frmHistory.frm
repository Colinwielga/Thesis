VERSION 5.00
Begin VB.Form frmHistory 
   BackColor       =   &H00FFFFFF&
   Caption         =   "History of Mario and Mario Madness Maker"
   ClientHeight    =   9945
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12375
   LinkTopic       =   "Form1"
   Picture         =   "frmHistory.frx":0000
   ScaleHeight     =   9945
   ScaleWidth      =   12375
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdHistory 
      BackColor       =   &H00FFFFFF&
      Caption         =   "The History of Mario"
      Height          =   852
      Left            =   6720
      TabIndex        =   4
      Top             =   7560
      Width           =   2292
   End
   Begin VB.CommandButton cmdworkscited 
      Caption         =   "Works Cited"
      Height          =   852
      Left            =   9240
      TabIndex        =   3
      Top             =   7560
      Width           =   2292
   End
   Begin VB.CommandButton cmdBill 
      Caption         =   "Information on Bill Macy, the Maker of Mario Madness"
      Height          =   852
      Left            =   6720
      TabIndex        =   2
      Top             =   8520
      Width           =   2292
   End
   Begin VB.PictureBox picoutput 
      Height          =   6132
      Left            =   5760
      ScaleHeight     =   6075
      ScaleWidth      =   5955
      TabIndex        =   1
      Top             =   960
      Width           =   6012
   End
   Begin VB.CommandButton cmdreturn 
      Caption         =   "Return to the main page"
      Height          =   852
      Left            =   9240
      TabIndex        =   0
      Top             =   8520
      Width           =   2292
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "By Bill Macy"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   9600
      Width           =   1935
   End
   Begin VB.Label lblInformation 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Mario Information!"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   25.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   972
      Left            =   3720
      TabIndex        =   5
      Top             =   240
      Width           =   6492
   End
End
Attribute VB_Name = "frmHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: Mario Madness
'Form name: frmHistory
'Author: Bill Macy
'Date Written: Tuesday March 14th, 2006
'Objective of form:  This form allows the user to view information on the history of Mario, along with
                'information on me, and all the sites that I referenced to make the project.  They can also
                'press a button to bring them back to the main page.  Everything is displayed in the picture box.

Option Explicit
Private Sub cmdBill_Click()
    picoutput.Cls       'clears the picture box of anything that was previously written in it
    picoutput.Print "The Maker"     'prints the words "The Maker" in the picture box
    picoutput.Print "***********************************"       'prints a bunch of stars
    picoutput.Print     'prints a blank line
    picoutput.Print "Bill Macy is currently a student at St John's University who is studying"      'prints the following information each line getting its own line in the picture box
    picoutput.Print "Management, Accounting, Economics, and Computer Science.  He is originally "
    picoutput.Print "from Maple Grove, Minnesota and loves to do anything outside including "
    picoutput.Print "going to his cabin in the summer.  Whenever he can, Bill tries his hardest"
    picoutput.Print "to learn something new.  Bill is extremely close to his family and has a"
    picoutput.Print "six year old beagle named Shadow.  Besides his mother and father, he has"
    picoutput.Print "one sibling - his sister, Kristi.  After he graduates from St John's, he"
    picoutput.Print "plans to pursure a career in whatever comes his way.  Live life for life"
    picoutput.Print "and not for money!"
End Sub


Private Sub cmdreturn_Click()
    frmHistory.Hide     'hides the history page
    frmMain.Show        'shows the main page
End Sub

Private Sub cmdworkscited_Click()
    picoutput.Cls       'clears the picture box of anything that was previously writtng in it
    picoutput.Print "Works Cited"       'prints the words "works cited"
    picoutput.Print "***********************************"       'prints a bunch of stars
    picoutput.Print "Wikipedia"     'prints the words "wikipedia"
    picoutput.Print Tab(30); "http://en.wikipedia.org/wiki/Mario"       'prints the corresponding web site
    picoutput.Print      'prints a blank line
    picoutput.Print "IGN"       'prints the letters "IGN"
    picoutput.Print Tab(30); "www.ign.com"      'prints the corresponding web site
    picoutput.Print     'prints a blank line
    picoutput.Print "Super Mario Bros. Headquarters"        'prints the words "Super Mario Bros. Headquarters"
    picoutput.Print Tab(30); "www.smbhq.com/"       'prints the corresponding web site
    picoutput.Print     'prints a blank line
    picoutput.Print "Mario Monsters"        'prints the words Mario Monsters
    picoutput.Print Tab(30); "www.mariomonsters.com"       'prints the corresponding web site
    picoutput.Print     'prints a blank line
    picoutput.Print "Free VB code (MarioCatcher)"       'prints the words "Free VB code (MarioCatcher)"
    picoutput.Print Tab(30); "www.freevbcode.com/ShowCode"      'prints the corresponding web site
    picoutput.Print     'prints a blank line
    picoutput.Print "Vb teaching tools"     'prints the words "Vb teaching tools"
    picoutput.Print Tab(30); "www.vbteachingtools.com"      'prints the corresponding web site
    picoutput.Print     'prints a blank line
    picoutput.Print "Nintendo Power"        'prints the words "Nintendo Power"
    picoutput.Print Tab(30); "www.nintendopower.com"        'prints the corresponding web site
    picoutput.Print     'prints a blank line
    picoutput.Print "Computer Concepts and Applications for Non-Majors (Our book)"      'prints the name of our book
    picoutput.Print     'prints a blank line
    picoutput.Print     'prints a blank line
    picoutput.Print     'prints a blank line
    picoutput.Print     'prints a blank line
    picoutput.Print "The pictures used in my program are from a wide variety of sites.  Many"  'prints the following information
    picoutput.Print "are from the pages I have already mentioned."
    
    
End Sub

Private Sub cmdHistory_Click()
    picoutput.Cls       'clears the picture box of anythign that may have previously been in it
    picoutput.Print "The History of Mario"      'prints the words "The history of Mario "
    picoutput.Print "***********************************"       'prints a bunch of stars
    picoutput.Print     'prints a blank line
    picoutput.Print "Super Mario Bros. is a video game produced by Nintendo in 1985. Universally"       'prints the following information
    picoutput.Print "considered a classic of the medium, Super Mario Bros. was one of the first "
    picoutput.Print "side-scrolling platform games of its kind, introducing players to huge, "
    picoutput.Print "bright, expansive worlds that changed the way video games were created."
    picoutput.Print "Super Mario Bros. is considered by The Guinness Book of World Records as "
    picoutput.Print "the best-selling video game of all time, and was largely responsible for "
    picoutput.Print "the initial success of the Famicom and Nintendo Entertainment System. It "
    picoutput.Print "has inspired countless imitators and was one of Shigeru Miyamoto's (creator)"
    picoutput.Print "most influential early successes. The game was turned into a film 2 months"
    picoutput.Print "later. The film was produced by Columbia Pictures and co-produced by"
    picoutput.Print "Nintendo Entertainment.  There was also a live-action 1993 version of the"
    picoutput.Print "film that was produced by Hollywood Pictures. The game, which starred Mario,"
    picoutput.Print "made him Nintendo's mascot, and who was at one time more recognizable"
    picoutput.Print "among American children than Mickey Mouse.  The game sold approximately"
    picoutput.Print "40 million copies worldwide which still stands as a Guinness World Record."
    picoutput.Print "It has been estimated that this game, next to Tetris, is the bestselling"
    picoutput.Print "game of all time. Although the game was popular enough on its own, mass"
    picoutput.Print "distribution is attributable to the popularity of the NES itself. Super"
    picoutput.Print "Mario Bros. was most often packaged with the console (usually in a dual"
    picoutput.Print "cartridge with the shooting game, Duck Hunt), just as Tetris was packaged"
    picoutput.Print "with the Game Boy. Super Mario Bros. 2 and 3 followed sometime after."
    picoutput.Print "Super Mario Bros.3 is often cited as the best selling non-packaged game of"
    picoutput.Print "all time."
End Sub
