VERSION 5.00
Begin VB.Form frmPeople 
   BackColor       =   &H80000007&
   Caption         =   "Legends"
   ClientHeight    =   6825
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10515
   LinkTopic       =   "Form1"
   ScaleHeight     =   6825
   ScaleWidth      =   10515
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to History"
      Height          =   615
      Left            =   8040
      TabIndex        =   5
      Top             =   5880
      Width           =   1935
   End
   Begin VB.PictureBox picresults1 
      Height          =   3735
      Left            =   3240
      ScaleHeight     =   3675
      ScaleWidth      =   2955
      TabIndex        =   4
      Top             =   600
      Width           =   3015
   End
   Begin VB.PictureBox picResults 
      Height          =   3735
      Left            =   6600
      ScaleHeight     =   3675
      ScaleWidth      =   3675
      TabIndex        =   3
      Top             =   600
      Width           =   3735
   End
   Begin VB.CommandButton cmdPerson3 
      Caption         =   "Bill France Sr."
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   2415
   End
   Begin VB.CommandButton cmdPerson2 
      Caption         =   "Dale Earnhardt Sr."
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   3600
      Width           =   2415
   End
   Begin VB.CommandButton cmdPerson1 
      Caption         =   "Richard Petty"
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   2040
      Width           =   2415
   End
End
Attribute VB_Name = "frmPeople"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Introduction to NASCAR
'Form People
'Colin Roberts and Luke Hommerding
'Written 10/18/09
'Purpose is to use dynamic picture loading to display legends of the Sport of Nascar
Option Explicit
'When first button is clicked info and picture are loaded in the picture boxes
Private Sub cmdPerson1_Click()
    picResults1.Cls
    picResults.Picture = LoadPicture(App.Path & "\richardpetty.jpg")
    picResults1.Print "- King of NASCAR"
    picResults1.Print
    picResults1.Print "- 200 Career Wins"
    picResults1.Print
    picResults1.Print "- 7 Winston Cup Championships"
    picResults1.Print
    picResults1.Print "- Racer and Friend to Thousands"
End Sub
'When second button is clicked info and picture are loaded in the picture boxes
Private Sub cmdPerson2_Click()
    picResults1.Cls
    picResults.Picture = LoadPicture(App.Path & "\earnhardt_dale1.jpg")
    picResults1.Print "- The Intimidator"
    picResults1.Print
    picResults1.Print "- Tragic Death provided new safety"
    picResults1.Print "  for future drivers (HANS DEVICE)"
    picResults1.Print
    picResults1.Print "- Attracted many new fans to the sport"
    picResults1.Print
    picResults1.Print "- Seven Winston Cup Championships"
End Sub
'When thrid button is clicked info and picture are loaded in the picture boxes
Private Sub cmdPerson3_Click()
    picResults1.Cls
    picResults.Picture = LoadPicture(App.Path & "\BillFrance.jpg")
    picResults1.Print "- Founder of NASCAR"
    picResults1.Print
    picResults1.Print "- Stressed the importance of"
    picResults1.Print "  keeping the business in the family"
    picResults1.Print
    picResults1.Print "- Used his racing experience to"
    picResults1.Print "  make the best organization he could"
    picResults1.Print
    picResults1.Print "- Inducted into Hall of Fame 4 times "
End Sub
'returns user to the history form
Private Sub cmdReturn_Click()
    frmHistory.Show
    frmPeople.Hide
End Sub
