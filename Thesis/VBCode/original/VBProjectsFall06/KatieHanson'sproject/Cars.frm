VERSION 5.00
Begin VB.Form Cars 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   7275
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10275
   LinkTopic       =   "Form1"
   ScaleHeight     =   7275
   ScaleWidth      =   10275
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picRating 
      BackColor       =   &H000000FF&
      Height          =   735
      Left            =   4920
      ScaleHeight     =   675
      ScaleWidth      =   1875
      TabIndex        =   5
      Top             =   4920
      Width           =   1935
   End
   Begin VB.CommandButton cmdRate 
      BackColor       =   &H000000FF&
      Caption         =   "Rate this Movie"
      Height          =   375
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5160
      Width           =   1575
   End
   Begin VB.CommandButton cmdPurchase 
      BackColor       =   &H000000FF&
      Caption         =   "Purchase Movie"
      Height          =   375
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6240
      Width           =   1935
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H000000FF&
      Caption         =   "Back To Title Page"
      Height          =   375
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6720
      Width           =   1935
   End
   Begin VB.PictureBox picCars1 
      Height          =   2535
      Left            =   5280
      Picture         =   "Cars.frx":0000
      ScaleHeight     =   2475
      ScaleWidth      =   3555
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
   Begin VB.Label lblName 
      BackColor       =   &H00000000&
      Caption         =   " Cars"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   38.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   1320
      TabIndex        =   8
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H000000FF&
      Caption         =   $"Cars.frx":4E8C
      BeginProperty Font 
         Name            =   "Maiandra GD"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   4215
   End
   Begin VB.Label lblVRating 
      BackColor       =   &H000000FF&
      Caption         =   "Viewers Rating:"
      Height          =   255
      Left            =   4920
      TabIndex        =   6
      Top             =   4680
      Width           =   1935
   End
   Begin VB.Label lblCRating 
      BackColor       =   &H00000000&
      Caption         =   "Critics Rating:          B"
      BeginProperty Font 
         Name            =   "Maiandra GD"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   5880
      TabIndex        =   3
      Top             =   3240
      Width           =   2415
   End
End
Attribute VB_Name = "Cars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Movies
'Form name: Cars
'Author: Katie Hanson
'Date Written: Nov 1 2006
'Objective: This form allows the user to get information about the movie "Cars". The user learns what the movie is about and also how the critics rated it. The user can also rate the movie and purchase it.

Option Explicit
'Back to the startup form
Private Sub cmdBack_Click()
    Title.Show
    Cars.Hide
End Sub
'Move to purchase form to purchase movie
Private Sub cmdPurchase_Click()
    Purchase.Show
    Cars.Hide
End Sub
'user rating movie
Private Sub cmdRate_Click()
Dim Rating As String
    Rating = InputBox("Share your feelings about this movie. Enter a letter (A,B,C,D) to share how you feel about this movie.")
    picRating.Print Rating
End Sub
