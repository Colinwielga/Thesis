VERSION 5.00
Begin VB.Form frmS 
   Caption         =   "Form1"
   ClientHeight    =   7650
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5745
   LinkTopic       =   "Form1"
   Picture         =   "frmS.frx":0000
   ScaleHeight     =   7650
   ScaleWidth      =   5745
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdKS 
      Caption         =   "Ko Simpson"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   2055
   End
   Begin VB.CommandButton cmdJA 
      Caption         =   "Jason Allen"
      Height          =   495
      Left            =   2640
      TabIndex        =   4
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton cmdDB 
      Caption         =   "Darnell Bing"
      Height          =   495
      Left            =   2880
      TabIndex        =   3
      Top             =   2640
      Width           =   2055
   End
   Begin VB.CommandButton cmdDW 
      Caption         =   "Donte Whitner"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   3240
      Width           =   2055
   End
   Begin VB.CommandButton cmdMH 
      Caption         =   "Michael Huff"
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "Back"
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   7080
      Width           =   1095
   End
End
Attribute VB_Name = "frmS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'show profiles via message boxes
Option Explicit
Private Sub cmdBack_Click()
    frmS.Hide
    frmAthletes.Show
End Sub

Private Sub cmdDB_Click()
    MsgBox "Positives: Darnell Bing has prototypical size for an NFL safety, with very, very good speed and hitting power to complement it. He is above-average in coverage, although not great in it. Bing has tons of upside, which is impressive given his current ability. He is not a bit afraid to come up to the line of scrimmage and blitz, and can lay the lumber on the ball carrier anywhere on the field.Negatives: Bing has never been able to put together a healthy season, seemingly always being hurt and missing some time. To his credit, he plays through a lot of injuries, but they effect his play more than they should. His coverage ability isn't what it could be with his speed and athleticism, but that may be in part because of those injuries.", , "Darnell Bing (USC)"
End Sub

Private Sub cmdDW_Click()
    MsgBox "Positives:One thing is for sure: whoever gets Dwayne Slay is getting a quarterback killer. He is a hard-hitting safety and a great tackler who plays more like a linebacker. He has great size for a free safety – the type of size (6-3, 214) that NFL GMs drool over. The best thing you can say about Slay is that he makes plays. In 2005, he posted a Texas Tech record eight forced fumbles to go along with his 101 tackles.Negatives:Slay’s coverage skills are largely unproven at the college level, and it shows in his career interception number of one. He is a bit slow for an NFL free safety and may have trouble keeping up with the speedsters he’ll have to stay on top of as a result. There are also serious questions about his experience, as he has only started for one season in college.", , "Donte Whitner (Ohio State)"
End Sub

Private Sub cmdJA_Click()
    MsgBox "Positives:There are three things about Allen that will really help him out in the future. The first being his intelligence. He's going to have to continue to be a smart player at the next level. The second good thing is that he is usually a sure tackler. His ability to get his head across and bring a player down may be his greatest asset. The last thing is that he has good size and speed. This will allow him to not be overmatched when he's in man coverage.  Negatives: He'll need to improve his ability to read a play just a little bit. If he doesn't learn to read better, he will be overmatched by some of the NFL's best QBs. He also needs to improve his route recognition some, or he may get fooled on routes run by the best receivers.", , "Jason Allen (Tennessee)"
End Sub

Private Sub cmdKS_Click()
    MsgBox "Positives:Kelly Jennings, one of the best three cornerbacks in an average 2006 class, is a fast, athletic cornerback who excels at playing off of opposing wide receivers. He also plays well in man-to-man coverage with most wideouts. He is known as a ballhawk-type of corner who plays the ball well in flight.Negatives: Jennings tends to have some problems with more physical receivers due to his slight frame. He will definitely be told to bulk up some by whatever team drafts him in April. Also, Jennings has a penchant for getting beat on out routes and curls because of the space he gives to receivers. When he can’t use his speed to catch up, it means trouble for his defense. Tackling and contact isn’t his strongest suit, and he’ll need to prove that he isn’t afraid of contact to excel at the next level.", , "Ko Simpson (South Carolina)"
End Sub

Private Sub cmdMH_Click()
    MsgBox "Positives:Michael Huff is a very athletic defensive back, who can play both corner and safety. He has excellent speed to go along with good hip movement which allows him to run with receivers and make plays on the ball. He's very smart player who recognizes routes. His athleticism and smarts make him a ballhawk, and he has a tendency to return interceptions for touchdowns.Negatives:Once in awhile he'll tend to gamble and try to make a big play, and instead he gets beat for a big gain or TD. He'll need to improve on his technique just a little bit to be pro ready. He's moved around his whole career at Texas between cornerback and safety, so he's a little bit of a 'tweener, but he should be able to make the adjustment to one position.", , "Michael Huff (Texas)"
End Sub
