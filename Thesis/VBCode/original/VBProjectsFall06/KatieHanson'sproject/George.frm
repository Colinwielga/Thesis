VERSION 5.00
Begin VB.Form George 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   6030
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9165
   LinkTopic       =   "Form1"
   ScaleHeight     =   6030
   ScaleWidth      =   9165
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRate 
      BackColor       =   &H0080FF80&
      Caption         =   "Rate this Movie"
      Height          =   495
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5040
      Width           =   1095
   End
   Begin VB.PictureBox picRating 
      BackColor       =   &H0080FF80&
      Height          =   975
      Left            =   3720
      ScaleHeight     =   915
      ScaleWidth      =   1515
      TabIndex        =   4
      Top             =   4800
      Width           =   1575
   End
   Begin VB.CommandButton cmdPurchase 
      BackColor       =   &H0080FF80&
      Caption         =   "Purchase Movie "
      Height          =   375
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5040
      Width           =   2055
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H0080FF80&
      Caption         =   "Back to Title Page"
      Height          =   375
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5520
      Width           =   2055
   End
   Begin VB.PictureBox picGeorge1 
      Height          =   2655
      Left            =   6120
      Picture         =   "George.frx":0000
      ScaleHeight     =   2595
      ScaleWidth      =   2715
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label lblName 
      BackColor       =   &H00000000&
      Caption         =   "Curious George"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   975
      Left            =   360
      TabIndex        =   8
      Top             =   240
      Width           =   3855
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H00C0FFC0&
      Caption         =   $"George.frx":29C2
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   3375
   End
   Begin VB.Label lblVRating 
      BackColor       =   &H0080FF80&
      Caption         =   "Viewers Rating:"
      Height          =   255
      Left            =   3720
      TabIndex        =   5
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label lblCRating 
      BackColor       =   &H00000000&
      Caption         =   "Critics Rating:        B-"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   1095
      Left            =   3960
      TabIndex        =   3
      Top             =   3000
      Width           =   2535
   End
End
Attribute VB_Name = "George"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Movies
'Form name: George
'Author: Katie Hanson
'Date Written: Nov 1 2006
'Objective: This form allows the user to get information about the movie "Curious George". The user learns what the movie is about and also how the critics rated it. The user can also rate the movie and purchase it.

Option Explicit
'Back to startup form
Private Sub cmdBack_Click()
    Title.Show
    George.Hide
End Sub
'move to purchase form to purchase movie
Private Sub cmdPurchase_Click()
    Purchase.Show
    George.Hide
End Sub
'user rating the movie
Private Sub cmdRate_Click()
Dim Rating As String
    Rating = InputBox("Share your feelings about this movie. Enter a letter (A,B,C,D) to share how you feel about this movie.")
    picRating.Print Rating
End Sub
