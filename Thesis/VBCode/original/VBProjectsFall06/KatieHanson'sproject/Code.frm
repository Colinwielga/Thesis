VERSION 5.00
Begin VB.Form Code 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   6585
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9765
   LinkTopic       =   "Form1"
   ScaleHeight     =   6585
   ScaleWidth      =   9765
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRate 
      Caption         =   "Rate this Movie"
      Height          =   495
      Left            =   6120
      TabIndex        =   7
      Top             =   5040
      Width           =   855
   End
   Begin VB.PictureBox picRating 
      Height          =   1335
      Left            =   4800
      ScaleHeight     =   1275
      ScaleWidth      =   1155
      TabIndex        =   5
      Top             =   4800
      Width           =   1215
   End
   Begin VB.PictureBox picCode2 
      Height          =   1455
      Left            =   7440
      Picture         =   "Code.frx":0000
      ScaleHeight     =   1395
      ScaleWidth      =   1875
      TabIndex        =   3
      Top             =   1920
      Width           =   1935
   End
   Begin VB.CommandButton cmdPurchase 
      Caption         =   "Purchase Movie"
      Height          =   375
      Left            =   7800
      TabIndex        =   2
      Top             =   5400
      Width           =   1695
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back To Title Page"
      Height          =   375
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6000
      Width           =   1695
   End
   Begin VB.PictureBox picCode1 
      Height          =   1335
      Left            =   7440
      Picture         =   "Code.frx":0D8E
      ScaleHeight     =   1275
      ScaleWidth      =   1875
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label lblInfo 
      Caption         =   $"Code.frx":1893
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   240
      TabIndex        =   9
      Top             =   1440
      Width           =   3855
   End
   Begin VB.Label lblName 
      BackColor       =   &H00000000&
      Caption         =   "The Davinci Code"
      BeginProperty Font 
         Name            =   "Gulim"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   615
      Left            =   240
      TabIndex        =   8
      Top             =   360
      Width           =   4815
   End
   Begin VB.Label lblVRating 
      Caption         =   "Viewers Rating:"
      Height          =   255
      Left            =   4800
      TabIndex        =   6
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label lblCRating 
      BackColor       =   &H00000000&
      Caption         =   "Critics Rating:          C+"
      BeginProperty Font 
         Name            =   "Gulim"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   855
      Left            =   4320
      TabIndex        =   4
      Top             =   2760
      Width           =   2775
   End
End
Attribute VB_Name = "Code"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Movies
'Form name: Code
'Author: Katie Hanson
'Date Written: Nov 1 2006
'Objective: This form allows the user to get information about the movie "The Davinci Code". The user learns what the movie is about and also how the critics rated it. The user can also rate the movie and purchase it.

Option Explicit
'Back to startup form
Private Sub cmdBack_Click()
    Title.Show
    Code.Hide
End Sub
'Move to purchase form to purchase movie
Private Sub cmdPurchase_Click()
    Purchase.Show
    Code.Hide
End Sub
'user rating the movie
Private Sub cmdRate_Click()
Dim Rating As String
    Rating = InputBox("Share your feelings about this movie. Enter a letter (A,B,C,D) to share how you feel about this movie.")
    picRating.Print Rating
End Sub
