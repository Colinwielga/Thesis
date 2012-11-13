VERSION 5.00
Begin VB.Form frmreadexcerpts 
   BackColor       =   &H0080C0FF&
   Caption         =   "Read Excerpts; Project by Kayla Nelson"
   ClientHeight    =   7530
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10200
   ForeColor       =   &H000040C0&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7530
   ScaleWidth      =   10200
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox piccity 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2565
      Left            =   8040
      Picture         =   "readexcerpts.frx":0000
      ScaleHeight     =   2535
      ScaleWidth      =   1500
      TabIndex        =   4
      Top             =   2040
      Width           =   1530
   End
   Begin VB.PictureBox pickite 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2490
      Left            =   5760
      Picture         =   "readexcerpts.frx":1CA1
      ScaleHeight     =   2460
      ScaleWidth      =   1575
      TabIndex        =   3
      Top             =   2040
      Width           =   1605
   End
   Begin VB.PictureBox picfirst 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   3120
      Picture         =   "readexcerpts.frx":A51A
      ScaleHeight     =   2625
      ScaleWidth      =   1785
      TabIndex        =   2
      Top             =   1920
      Width           =   1815
   End
   Begin VB.PictureBox picmillion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   480
      Picture         =   "readexcerpts.frx":C188
      ScaleHeight     =   2985
      ScaleWidth      =   1905
      TabIndex        =   1
      Top             =   1800
      Width           =   1935
   End
   Begin VB.CommandButton cmdreturnmm 
      BackColor       =   &H000080FF&
      Caption         =   "Return to Main Menu"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6600
      Width           =   2775
   End
   Begin VB.Label lblreadexcerpt 
      BackColor       =   &H0080C0FF&
      Caption         =   "Click on the book you would like to read an excerpt on."
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   5
      Top             =   120
      Width           =   9495
   End
End
Attribute VB_Name = "frmreadexcerpts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Kayla's Book Club (MainMenu.vbp)
'Form Name: Read Excerpts (frmreadexcerpts.frm)
'Author: Kayla Nelson
'purpose of the form: This form has the user click on the book they would like to read an excerpt on.  WHen they click on the book, a Message box will appear with the excerpt from that book.


Option Explicit
Private Sub cmdreturnmm_Click() 'This closes the Read Excerpt form and brings the user back to the Main Menu
    frmreadexcerpts.Hide
    frmmainmenu.Show
End Sub

Private Sub piccity_Click()
    MsgBox "The air still smelled of charcoal when I arrived in Venice three days after the fire. As it happened, the timing of my visit was purely coincidental. I had made plans, months before, to come to Venice for a few weeks in the off-season in order to enjoy the city without the crush of other tourists. If there had been a wind Monday night, the water-taxi driver told me as we came across the lagoon from the airport, there wouldn't be a Venice to come to. How did it happen? I asked.The taxi driver shrugged. How do all these things happen?", , "Excerpt from The City of Fallen Angels" 'A message box appears with the following excerpt
End Sub

Private Sub picfirst_Click()
    MsgBox "See, it’s simple,” Alvin said. “First, you meet a nice girl, and then you date for a while to make sure you share the same values. See if you two are compatible in the big, ‘this is our life and we’re in it together’ decisions. You know, talk about which family you’re going to visit on the holidays, whether you want to live in a house or an apartment, whether to get a dog or a cat, who gets to use the shower first in the morning, while there’s still plenty of hot water. If you two are still pretty much in agreement, then you get married. Are you following me here?” “I’m following you,” Jeremy said.", , "Excerpt from At First Sight" 'A message box appears from the following excerpt
End Sub

Private Sub pickite_Click()
    MsgBox "I became what I am today at the age of twelve, on a frigid overcast day in the winter of 1975. I remember the precise moment, crouching behind a crumbling mud wall, peeking into the alley near the frozen creek. That was a long time ago, but it’s wrong what they say about the past, I’ve learned, about how you can bury it. Because the past claws its way out. Looking back now, I realize I have been peeking into that deserted alley for the last twenty-six years.", , "Excerpt from The Kite Runner" 'A Message box appears with the following excerpt
End Sub

Private Sub picmillion_Click()
    MsgBox "I walk toward a door where a Nurse stands waiting for me. As I walk past her she is careful not to touch me and I am brought back from the happy afterglow of pachyderm memories and I am reminded of what I am. I am an Alcoholic and I am a drug Addict and I am a Criminal. I am missing my front four teeth. I have a hole in my cheek that has been closed with forty-one stitches. I have a broken nose and I have black swollen eyes. I have an Escort because I am a Patient at a Drug and Alcohol Treatment Center. I am wearing a borrowed jacket because I don't have one of my own. I am carrying two old yellow tennis balls because I'm not allowed to have any painkillers or anesthesia. I am an Alcoholic. I am a drug Addict. I am a Criminal. ", , "Excerpt from A Million Little Pieces" 'A Message box appears with the following excerpt.
End Sub
