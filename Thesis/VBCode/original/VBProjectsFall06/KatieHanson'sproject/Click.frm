VERSION 5.00
Begin VB.Form Click 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   6090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9705
   LinkTopic       =   "Form1"
   ScaleHeight     =   6090
   ScaleWidth      =   9705
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRate 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Rate this Movie"
      Height          =   495
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4080
      Width           =   1215
   End
   Begin VB.PictureBox picRating 
      BackColor       =   &H00C0FFFF&
      Height          =   1335
      Left            =   4200
      ScaleHeight     =   1275
      ScaleWidth      =   1395
      TabIndex        =   4
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton cmdpurchase 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Purchase Movie"
      Height          =   375
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5160
      Width           =   1575
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Back to Title Page"
      Height          =   375
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5640
      Width           =   1575
   End
   Begin VB.PictureBox picClick1 
      BackColor       =   &H00000000&
      Height          =   3015
      Left            =   6120
      Picture         =   "Click.frx":0000
      ScaleHeight     =   2955
      ScaleWidth      =   3195
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label lblName 
      BackColor       =   &H00000000&
      Caption         =   "Click"
      BeginProperty Font 
         Name            =   "Eras Light ITC"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   975
      Left            =   1200
      TabIndex        =   8
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H00C0FFFF&
      Caption         =   $"Click.frx":3258
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   3375
   End
   Begin VB.Label lblVRating 
      BackColor       =   &H00C0FFFF&
      Caption         =   "viewer rating:"
      Height          =   255
      Left            =   4200
      TabIndex        =   5
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label lbClRating 
      BackColor       =   &H00000000&
      Caption         =   "Critics Rating:          C"
      BeginProperty Font 
         Name            =   "Eras Light ITC"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   855
      Left            =   3720
      TabIndex        =   3
      Top             =   2400
      Width           =   2295
   End
End
Attribute VB_Name = "Click"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Movies
'Form name: Click
'Author: Katie Hanson
'Date Written: Nov 1 2006
'Objective: This form allows the user to get information about the movie "Click". The user learns what the movie is about and also how the critics rated it. The user can also rate the movie and purchase it.

Option Explicit
'Back to startup form
Private Sub cmdBack_Click()
    Title.Show
    Click.Hide
End Sub
'Move to purcahse form to purchase movie
Private Sub cmdPurchase_Click()
    Purchase.Show
    Click.Hide
End Sub
'user rating movie
Private Sub cmdRate_Click()
Dim Rating As String
    Rating = InputBox("Share your feelings about this movie. Enter a letter (A,B,C,D) to share how you feel about this movie.")
    picRating.Print Rating
End Sub


