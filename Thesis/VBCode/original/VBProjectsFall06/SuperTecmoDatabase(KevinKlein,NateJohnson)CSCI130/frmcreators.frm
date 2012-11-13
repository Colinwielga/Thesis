VERSION 5.00
Begin VB.Form frmcreators 
   BackColor       =   &H00400000&
   Caption         =   "About the Creators"
   ClientHeight    =   7440
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10140
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   10140
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picresults3 
      Height          =   3495
      Left            =   3480
      ScaleHeight     =   3435
      ScaleWidth      =   4035
      TabIndex        =   5
      Top             =   3840
      Width           =   4095
   End
   Begin VB.PictureBox picresults2 
      Height          =   3495
      Left            =   120
      ScaleHeight     =   3435
      ScaleWidth      =   3195
      TabIndex        =   4
      Top             =   3840
      Width           =   3255
   End
   Begin VB.PictureBox picresults 
      Height          =   3615
      Left            =   2640
      ScaleHeight     =   3555
      ScaleWidth      =   6195
      TabIndex        =   3
      Top             =   120
      Width           =   6255
   End
   Begin VB.CommandButton cmdback 
      BackColor       =   &H000000FF&
      Caption         =   "Go Back"
      Height          =   3495
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3840
      Width           =   2415
   End
   Begin VB.CommandButton cmdkevin 
      BackColor       =   &H000000FF&
      Caption         =   "Kevin"
      Height          =   1695
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2040
      Width           =   2295
   End
   Begin VB.CommandButton cmdnate 
      BackColor       =   &H000000FF&
      Caption         =   "Nate"
      Height          =   1815
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmcreators"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: Super Tecmo Database
'Form name: frmcreators
'Author: Nate Johnson & Kevin Klein
'Date Written: October 11th, 2006
'Objective of project: This project will allow its users to learn more about the game of football
'and will also allow them the oppurtunity to learn how to play the game of football with the Nintendo
'video game, Tecmo Super Bowl.
'Objective of form: This form allows the users to see who was responsible for creating
'this project, it also allows the user to see our pictures and to learn some background
' information about the creators.

Private Sub CmdBack_Click()
frmcreators.Hide 'hides creators form'
frmHistory.Show 'shows history form'
End Sub

Private Sub cmdkevin_Click()
picresults.Cls 'clears the picbox of any content'
picresults2.Cls 'clears the picbox of any content'
picresults3.Cls 'clears the picbox of any content'
picresults2.Picture = LoadPicture(App.Path & "\pics\kevin1.jpg")
picresults3.Picture = LoadPicture(App.Path & "\pics\kevin2.jpg")
picresults.Print "Hi, my name is Kevin Klein and I am 21 years old." 'displays biography information'
picresults.Print "I went to Minnetonka high school and graduated"
picresults.Print "in 2004.  My major here at St. John's is business"
picresults.Print "management.  The reason I wanted to do this "
picresults.Print "Computer Science project was because I am a nerd"
picresults.Print "at heart.  Even though my major does not indicate"
picresults.Print "that I may be a nerd, I have always loved video "
picresults.Print "games.  I have been playing video games from the"
picresults.Print "Atari video system to the Sega Genesis, and the "
picresults.Print "Super Nintendo all the way to the X Box and the"
picresults.Print "amazing X Box 360.  So when my partner asked me "
picresults.Print "if I wanted to do a program about Tecmo Super Bowl"
picresults.Print "I was definitely excited.  I was also excited in "
picresults.Print "the fact that I never owned Tecmo Super Bowl, so "
picresults.Print "this gave me the chance to play and learn about "
picresults.Print "the game, and realize how much of a sweet game it really is."
End Sub

Private Sub cmdnate_Click()
picresults.Cls 'clears the picbox of any content'
picresults2.Cls 'clears the picbox of any content'
picresults3.Cls 'clears the picbox of any content'
picresults.Print "Hi My name is Nate Johnson and I am 20 years old and I am" 'displays biography information'
picresults.Print "a junior at Saint John's. I graduated from Duluth Denfeld"
picresults.Print "High School in 2004. My major here at SJU is Management with"
picresults.Print "a History minor. I wanted to do this project because of my"
picresults.Print "love for video games, especially older ones, and because "
picresults.Print "of my previous experience with this game. My cousin owned"
picresults.Print "this game and always brought it with him when he came to"
picresults.Print "baby-sit me and my brother.I thought that it was an"
picresults.Print "unbelieveable game, and I still have trouble playing other "
picresults.Print "football games because this one is so much fun. I have logged"
picresults.Print "hundred of hours on this game and thought that it would be "
picresults.Print "really interesting and fun to create a sort-of memorial of "
picresults.Print "information and experiences to this game."
picresults2.Picture = LoadPicture(App.Path & "\pics\nate3.bmp")
picresults3.Picture = LoadPicture(App.Path & "\pics\nate2.jpg") 'displays new picture
End Sub


