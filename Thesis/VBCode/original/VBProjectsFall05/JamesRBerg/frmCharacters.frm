VERSION 5.00
Begin VB.Form frmCharacters 
   Caption         =   "Characters"
   ClientHeight    =   8805
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10785
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   15.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "frmCharacters.frx":0000
   ScaleHeight     =   8805
   ScaleWidth      =   10785
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "Return"
      Height          =   855
      Left            =   720
      TabIndex        =   5
      Top             =   1200
      Width           =   1335
   End
   Begin VB.PictureBox Picture5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   8280
      Picture         =   "frmCharacters.frx":8274
      ScaleHeight     =   2835
      ScaleWidth      =   2115
      TabIndex        =   4
      Top             =   5760
      Width           =   2175
   End
   Begin VB.PictureBox Picture4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   3960
      Picture         =   "frmCharacters.frx":8845
      ScaleHeight     =   4395
      ScaleWidth      =   3435
      TabIndex        =   3
      Top             =   4320
      Width           =   3495
   End
   Begin VB.PictureBox Picture3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   840
      Picture         =   "frmCharacters.frx":CADA
      ScaleHeight     =   3675
      ScaleWidth      =   1995
      TabIndex        =   2
      Top             =   4920
      Width           =   2055
   End
   Begin VB.PictureBox Picture2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   4200
      Picture         =   "frmCharacters.frx":F9B2
      ScaleHeight     =   3795
      ScaleWidth      =   2235
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   7800
      Picture         =   "frmCharacters.frx":10E01
      ScaleHeight     =   5355
      ScaleWidth      =   2715
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label lblInstructions 
      BackColor       =   &H80000000&
      Caption         =   "To learn about each member of the Simpson family click on their picture!"
      Height          =   1935
      Left            =   360
      TabIndex        =   6
      Top             =   2400
      Width           =   2775
   End
End
Attribute VB_Name = "frmCharacters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Simpsons TV Show test (final.vbp)
'main form (highscore.frm)
'Jim Berg
'October 30, 2005
'This form shows the Simpson family and gives a short bio and most of the info that the test is created with

Option Explicit


Private Sub cmdPrevious_Click()
    frmmain.Show
    frmCharacters.Hide
End Sub

Private Sub Picture1_Click()
    MsgBox "Marge Simpson is a happy homemaker and mother of three. Her prides and joys are Bart , Lisa and Maggie. She's also very proud of her husband, Homer, even though he frequently loses his keys and needs her to find them. Marge also has strong relationships with her sisters, Patty and Selma, and with her father-in-law, Abe Simpson. Aside from her duties at home, Marge has flirted briefly with a number of careers ranging from police officer to anti-violence activist.", , "Marge's Bio"
End Sub

Private Sub Picture2_Click()
    MsgBox "Bart Simpson is misunderstood. Wrongly pegged as an underachiever and troublemaker, Bart would like to remind the world of some of his decent qualities: He looks out for his sister, Lisa; he's befriended outcasts and misfits like his best friend Milhouse Van Houten and Ralph Wiggum.  At age 10, Bart has managed to live out a number of dreams: He has starred in his own short-lived TV series (with his idol, Krusty the Clown), spotted, named a deadly comet that nearly destroyed his town and has avoided many killing attempts by Sideshow Bob. He couldn't have done any of those things without the help and support of his dog, Santa's Little Helper.", , "Barts Bio"
End Sub

Private Sub Picture3_Click()
    MsgBox "Homer juggles the roles of husband, father, safety inspector at the Springfield Nuclear Power Plant, bowler, beer drinker, astronaut, and dreamer, and makes it all look easy. Together with his high school sweetheart, Marge Bouvier, Homer settled down in Evergreen Terrace. He is fond of Duff Beer from Moes bar, donuts, Marge's pork chops and watching the Bee Guy on the Spanish channel.", , "Homer's Bio"
End Sub

Private Sub Picture4_Click()
    MsgBox "Lisa Simpson can't wait for college. She's only 8 and already reads at a 14th grade level, and has written a number of application-quality essays, one of which won her family a free trip to Washington, D.C. Her favorite activities include playing her saxophone, attending school and reading Non-Threatening Boys Magazine. A fan of Malibu Stacy, Lisa tried unsuccessfully to create her own talking doll, Lisa Lionheart. Unfortunately, no one wanted to buy a talking doll that was as judgmental as Lisa. Lisa wants everyone to know that she is a vegetarian and that if she could have one thing, it would be a pony, although she is fine with her cat Snowball", , "Lisa's Bio"
End Sub

Private Sub Picture5_Click()
    MsgBox "Maggie Simpson has done a lot in her one year of life. She's learned to spell her own name with an Etch A Sketch, she's wandered the town of Springfield all by herself, and she's shot Springfield's richest man because he attempted to steal her lollipop. Eventually, she'd like to learn how to speak and walk without falling down.", , "Maggie's Bio"
End Sub
