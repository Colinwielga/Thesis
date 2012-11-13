VERSION 5.00
Begin VB.Form breakup 
   BackColor       =   &H00000000&
   Caption         =   "Form2"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9360
   LinkTopic       =   "Form2"
   ScaleHeight     =   6795
   ScaleWidth      =   9360
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRating 
      BackColor       =   &H008080FF&
      Caption         =   "Rate this Movie"
      Height          =   375
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4440
      Width           =   1335
   End
   Begin VB.PictureBox picRating 
      BackColor       =   &H008080FF&
      Height          =   615
      Left            =   5160
      ScaleHeight     =   555
      ScaleWidth      =   2115
      TabIndex        =   4
      Top             =   4320
      Width           =   2175
   End
   Begin VB.CommandButton cmdPurchase 
      BackColor       =   &H008080FF&
      Caption         =   "Purchase Movie"
      Height          =   495
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5640
      Width           =   2055
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H008080FF&
      Caption         =   "Back To Title Page"
      Height          =   495
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6240
      Width           =   2055
   End
   Begin VB.PictureBox picbreakup2 
      Height          =   1935
      Left            =   6000
      Picture         =   "Form2.frx":0000
      ScaleHeight     =   1875
      ScaleWidth      =   2955
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label lblName 
      BackColor       =   &H00000000&
      Caption         =   " The Break Up"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   855
      Left            =   120
      TabIndex        =   8
      Top             =   240
      Width           =   4575
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H008080FF&
      Caption         =   $"Form2.frx":353E
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   4815
   End
   Begin VB.Label lblVRating 
      BackColor       =   &H008080FF&
      Caption         =   "Viewers Rating:"
      Height          =   255
      Left            =   5160
      TabIndex        =   5
      Top             =   4080
      Width           =   2175
   End
   Begin VB.Label lblCRating 
      BackColor       =   &H00000000&
      Caption         =   "Critics Rating:            C"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   855
      Left            =   6240
      TabIndex        =   3
      Top             =   2520
      Width           =   2175
   End
End
Attribute VB_Name = "breakup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Movies
'Form name: breakup
'Author: Katie Hanson
'Date Written: Nov 1 2006
'Objective: This form allows the user to get information about the movie "The Breakup". The user learns what the movie is about and also how the critics rated it. The user can also rate the movie and purchase it.

Option Explicit
'Back to the startup form
Private Sub cmdBack_Click()
    Title.Show
    breakup.Hide
End Sub
'Move to the purchase form to purchase the movie
Private Sub cmdPurchase_Click()
    Purchase.Show
    breakup.Hide
End Sub
'User rate the movie
Private Sub cmdRating_Click()
Dim Rating As String
    Rating = InputBox("Share your feelings about this movie. Enter a letter (A,B,C,D) to share how you feel about this movie.")
    picRating.Print Rating
End Sub
