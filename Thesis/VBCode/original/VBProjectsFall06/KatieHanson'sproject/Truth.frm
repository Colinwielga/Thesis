VERSION 5.00
Begin VB.Form Truth 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   6720
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10710
   LinkTopic       =   "Form1"
   ScaleHeight     =   6720
   ScaleWidth      =   10710
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRate 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Rate this Movie"
      Height          =   495
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4560
      Width           =   1455
   End
   Begin VB.PictureBox picRating 
      BackColor       =   &H00FFFFC0&
      Height          =   1335
      Left            =   5400
      ScaleHeight     =   1275
      ScaleWidth      =   1395
      TabIndex        =   4
      Top             =   4800
      Width           =   1455
   End
   Begin VB.PictureBox picTruth1 
      Height          =   2775
      Left            =   6480
      Picture         =   "Truth.frx":0000
      ScaleHeight     =   2715
      ScaleWidth      =   3795
      TabIndex        =   2
      Top             =   240
      Width           =   3855
   End
   Begin VB.CommandButton cmdPurchase 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Purchase Movie"
      Height          =   375
      Left            =   8280
      MaskColor       =   &H00FFFFC0&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5520
      Width           =   2055
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Back To Title Page"
      Height          =   375
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6120
      Width           =   2055
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H00FFFFC0&
      Caption         =   $"Truth.frx":1968
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   480
      TabIndex        =   8
      Top             =   2040
      Width           =   3735
   End
   Begin VB.Label lblName 
      BackColor       =   &H00000000&
      Caption         =   "An Inconvenient Truth"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   735
      Left            =   240
      TabIndex        =   7
      Top             =   240
      Width           =   5655
   End
   Begin VB.Label lblVRating 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Viewers Rating:"
      Height          =   255
      Left            =   5400
      TabIndex        =   5
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label lblCRating 
      BackColor       =   &H00000000&
      Caption         =   "Critics Rating:           B+"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   855
      Left            =   4920
      TabIndex        =   3
      Top             =   3360
      Width           =   2415
   End
End
Attribute VB_Name = "Truth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Movies
'Form name: Truth
'Author: Katie Hanson
'Date Written: Nov 1 2006
'Objective: This form allows the user to get information about the movie "An Inconvenient Truth". The user learns what the movie is about and also how the critics rated it. The user can also rate the movie and purchase it.

Option Explicit
'back to startup form
Private Sub cmdBack_Click()
    Title.Show
    Truth.Hide
End Sub
'move to purchase form to purchase movie
Private Sub cmdPurchase_Click()
    Purchase.Show
    Truth.Hide
End Sub
'user rating the movie
Private Sub cmdRate_Click()
Dim Rating As String
    Rating = InputBox("Share your feelings about this movie. Enter a letter (A,B,C,D) to share how you feel about this movie.")
    picRating.Print Rating
End Sub
