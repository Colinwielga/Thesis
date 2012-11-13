VERSION 5.00
Begin VB.Form Pirates 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   7020
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10395
   LinkTopic       =   "Form1"
   ScaleHeight     =   7020
   ScaleWidth      =   10395
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRate 
      BackColor       =   &H00000080&
      Caption         =   "Rate this Movie"
      Height          =   615
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5160
      Width           =   855
   End
   Begin VB.PictureBox picRating 
      BackColor       =   &H00000080&
      Height          =   1335
      Left            =   5280
      ScaleHeight     =   1275
      ScaleWidth      =   1635
      TabIndex        =   4
      Top             =   3480
      Width           =   1695
   End
   Begin VB.CommandButton cmdPurchase 
      BackColor       =   &H00000080&
      Caption         =   "Purchase Movie"
      Height          =   375
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5880
      Width           =   2055
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00000080&
      Caption         =   "Back to Title Page"
      Height          =   375
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6360
      Width           =   2055
   End
   Begin VB.PictureBox picPirates1 
      Height          =   2535
      Left            =   5880
      Picture         =   "Pirates.frx":0000
      ScaleHeight     =   2475
      ScaleWidth      =   3555
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H00000080&
      Caption         =   $"Pirates.frx":2B51
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   360
      TabIndex        =   8
      Top             =   1320
      Width           =   4695
   End
   Begin VB.Label lblName 
      BackColor       =   &H00000000&
      Caption         =   "The Pirates of the Caribbean:          Dead Man's Chest"
      BeginProperty Font 
         Name            =   "Juice ITC"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1095
      Left            =   600
      TabIndex        =   7
      Top             =   120
      Width           =   4575
   End
   Begin VB.Label lblVRating 
      BackColor       =   &H00000080&
      Caption         =   "Viewers Rating:"
      Height          =   255
      Left            =   5280
      TabIndex        =   5
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label lblCRating 
      BackColor       =   &H00000000&
      Caption         =   "Critics Rating:           B-"
      BeginProperty Font 
         Name            =   "Juice ITC"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   855
      Left            =   7560
      TabIndex        =   3
      Top             =   3600
      Width           =   1815
   End
End
Attribute VB_Name = "Pirates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Movies
'Form name: Pirates
'Author: Katie Hanson
'Date Written: Nov 1 2006
'Objective: This form allows the user to get information about the movie "The Pirates of the Caribbean". The user learns what the movie is about and also how the critics rated it. The user can also rate the movie and purchase it.

Option Explicit
'back to startup form
Private Sub cmdBack_Click()
    Title.Show
    Pirates.Hide
End Sub
'move to purchase form to purchase movie
Private Sub cmdPurchase_Click()
    Purchase.Show
    Pirates.Hide
End Sub
'user rating the movie
Private Sub cmdRate_Click()
Dim Rating As String
    Rating = InputBox("Share your feelings about this movie. Enter a letter (A,B,C,D) to share how you feel about this movie.")
    picRating.Print Rating
End Sub

