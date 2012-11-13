VERSION 5.00
Begin VB.Form frmShaunWhite 
   BackColor       =   &H000000C0&
   Caption         =   "Shaun White"
   ClientHeight    =   7965
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9045
   LinkTopic       =   "Form1"
   ScaleHeight     =   7965
   ScaleWidth      =   9045
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRange 
      Caption         =   "Shaun's Range of Scores"
      Height          =   495
      Left            =   3240
      TabIndex        =   10
      Top             =   6600
      Width           =   2775
   End
   Begin VB.CommandButton cmdMedals 
      Caption         =   "Shaun's Medal History"
      Height          =   495
      Left            =   3240
      TabIndex        =   9
      Top             =   5880
      Width           =   2775
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About Shaun White"
      Height          =   495
      Left            =   3240
      TabIndex        =   8
      Top             =   5160
      Width           =   2775
   End
   Begin VB.CommandButton cmdSlopeDescription 
      Caption         =   "Slopestlye Description"
      Height          =   855
      Left            =   720
      TabIndex        =   7
      Top             =   1320
      Width           =   1935
   End
   Begin VB.CommandButton cmdSuperDescription 
      Caption         =   "Superpipe Description"
      Height          =   855
      Left            =   6480
      TabIndex        =   6
      Top             =   1320
      Width           =   1935
   End
   Begin VB.CommandButton cmdPicture 
      Caption         =   "See picture and biography"
      Height          =   495
      Left            =   3120
      TabIndex        =   5
      Top             =   7320
      Width           =   1575
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   615
      Left            =   7080
      TabIndex        =   3
      Top             =   7080
      Width           =   1455
   End
   Begin VB.CommandButton cmdSuperpipe 
      Caption         =   "Take me to superpipe!"
      Height          =   855
      Left            =   6480
      TabIndex        =   1
      Top             =   3360
      Width           =   1935
   End
   Begin VB.CommandButton cmdSlopestyle 
      Caption         =   "Take me to slopestyle!"
      Height          =   855
      Left            =   720
      TabIndex        =   0
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Label lblCreator 
      BackColor       =   &H000000C0&
      Caption         =   "Creator: Andrew Dealy"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   7320
      Width           =   2895
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Shaun White's Winter X-Games History"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   555
      Left            =   480
      TabIndex        =   2
      Top             =   120
      Width           =   8145
   End
   Begin VB.Image Image1 
      Height          =   4065
      Left            =   3240
      Picture         =   "frmShaunWhite.frx":0000
      Top             =   960
      Width           =   2700
   End
End
Attribute VB_Name = "frmShaunWhite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This program describes and shows data on the professional snowboarder Shaun White's career
'This main page offers links to a varity of pages. There are several pages to descriptions
'of the creator, Shaun White, and different snowboarding competitions. This page also
'has links to the two types of competitions, slopestyle and superpipe. The page also
'includes a link to Shaun's medal history and a page of his range of scores.
'The program allows the user complete freedom to explore the world of Shaun White
'and his history in the winter x-games.
Private Sub cmdAbout_Click()
    frmBiography.Show
    frmShaunWhite.Hide
End Sub
'exits the program
Private Sub cmdExit_Click()
End
End Sub
'brings user to the from Medals
Private Sub cmdMedals_Click()
    frmMedals.Show
    frmShaunWhite.Hide
End Sub
'brings user to form Picture
Private Sub cmdPicture_Click()
 frmPicture.Show
 frmShaunWhite.Hide
End Sub
'brings user to form Range
Private Sub cmdRange_Click()
    frmRange.Show
    frmShaunWhite.Hide
End Sub
'brings user to form slope description
Private Sub cmdSlopeDescription_Click()
    frmSlopeDescription.Show
    frmShaunWhite.Hide
End Sub
'brings user to form slopestyle
Private Sub cmdSlopestyle_Click()
    frmSlopestyle.Show
    frmShaunWhite.Hide
End Sub
'brings user to form superpipe description
Private Sub cmdSuperDescription_Click()
    frmSuperDescription.Show
    frmShaunWhite.Hide
End Sub
'brings user to form superpipe
Private Sub cmdSuperpipe_Click()
    frmSuperpipe.Show
    frmShaunWhite.Hide
End Sub
