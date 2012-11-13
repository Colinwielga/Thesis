VERSION 5.00
Begin VB.Form lakehouse 
   BackColor       =   &H00000000&
   Caption         =   "Form3"
   ClientHeight    =   6585
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9405
   LinkTopic       =   "Form3"
   ScaleHeight     =   6585
   ScaleWidth      =   9405
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRate 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Rate this Movie"
      Height          =   495
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5400
      Width           =   1335
   End
   Begin VB.PictureBox picRating 
      BackColor       =   &H00FFFFC0&
      Height          =   975
      Left            =   3960
      ScaleHeight     =   915
      ScaleWidth      =   1755
      TabIndex        =   4
      Top             =   4200
      Width           =   1815
   End
   Begin VB.CommandButton cmdPurchase 
      BackColor       =   &H00FFFFC0&
      Cancel          =   -1  'True
      Caption         =   "Purchase Movie"
      Height          =   375
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5520
      Width           =   1935
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Back To Title Page"
      Height          =   375
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6000
      Width           =   1935
   End
   Begin VB.PictureBox piclakehouse1 
      Height          =   3495
      Left            =   3960
      Picture         =   "lakehouse.frx":0000
      ScaleHeight     =   3435
      ScaleWidth      =   5235
      TabIndex        =   0
      Top             =   120
      Width           =   5295
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H00FFFFC0&
      Caption         =   $"lakehouse.frx":6F05
      BeginProperty Font 
         Name            =   "High Tower Text"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   360
      TabIndex        =   8
      Top             =   1920
      Width           =   3135
   End
   Begin VB.Label lblName 
      BackColor       =   &H00000000&
      Caption         =   "The Lake     House"
      BeginProperty Font 
         Name            =   "High Tower Text"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   1215
      Left            =   840
      TabIndex        =   7
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label lblVRating 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Viewers Rating:"
      Height          =   255
      Left            =   3960
      TabIndex        =   5
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label lblCRating 
      BackColor       =   &H00000000&
      Caption         =   "Critics Rating:           C+"
      BeginProperty Font 
         Name            =   "High Tower Text"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   855
      Left            =   6720
      TabIndex        =   3
      Top             =   3720
      Width           =   2415
   End
End
Attribute VB_Name = "lakehouse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Movies
'Form name: lakehouse
'Author: Katie Hanson
'Date Written: Nov 1 2006
'Objective: This form allows the user to get information about the movie "The Lake House". The user learns what the movie is about and also how the critics rated it. The user can also rate the movie and purchase it.

Option Explicit
'back startup form
Private Sub cmdBack_Click()
    Title.Show
    lakehouse.Hide
End Sub
'move to purchase form to purchase movie
Private Sub cmdPurchase_Click()
    Purchase.Show
    lakehouse.Hide
End Sub
'user rating the movie
Private Sub cmdRate_Click()
Dim Rating As String
    Rating = InputBox("Share your feelings about this movie. Enter a letter (A,B,C,D) to share how you feel about this movie.")
    picRating.Print Rating
End Sub
