VERSION 5.00
Begin VB.Form Activites 
   BackColor       =   &H80000002&
   Caption         =   "Activities"
   ClientHeight    =   8760
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10995
   LinkTopic       =   "Form3"
   ScaleHeight     =   8760
   ScaleWidth      =   10995
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdboat 
      BackColor       =   &H0000FFFF&
      Caption         =   "Went on a Pirate Ship"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6600
      Width           =   2175
   End
   Begin VB.CommandButton cmdboom 
      BackColor       =   &H000000FF&
      Caption         =   "Made a Boomerang"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5520
      Width           =   2175
   End
   Begin VB.CommandButton cmddance 
      BackColor       =   &H0000FFFF&
      Caption         =   "Went to a Dance"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8160
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4440
      Width           =   2175
   End
   Begin VB.CommandButton cmdeat 
      BackColor       =   &H000000FF&
      Caption         =   "Ate Lots of Ice Cream"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3360
      Width           =   2175
   End
   Begin VB.PictureBox picresults2 
      BackColor       =   &H80000002&
      Height          =   6375
      Left            =   240
      ScaleHeight     =   6315
      ScaleWidth      =   7155
      TabIndex        =   5
      Top             =   1080
      Width           =   7215
   End
   Begin VB.CommandButton cmdback 
      BackColor       =   &H000000FF&
      Caption         =   "Back to Home Page"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7680
      Width           =   2175
   End
   Begin VB.CommandButton cmdclass 
      BackColor       =   &H0000FFFF&
      Caption         =   "Class"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2280
      Width           =   2175
   End
   Begin VB.CommandButton cmdpub 
      BackColor       =   &H000000FF&
      Caption         =   "Pub"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1200
      Width           =   2175
   End
   Begin VB.CommandButton cmdbeach 
      BackColor       =   &H0000FFFF&
      Caption         =   "Beach"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8160
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000002&
      Caption         =   "Ashley K. Smithson"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0FF&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   8160
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Oh The Places We Go"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "Activites"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Australia
'Form Name: Activites
'Author: Ashley Smithson
'Date: October 31, 2005
'Purpose of Form: to show all the fun stuff you can do!
Option Explicit
Private Sub cmdback_Click()
Activites.Hide 'brings you back to the main page
FinalProject2.Show
End Sub

Private Sub cmdbeach_Click()
picresults2.Picture = LoadPicture(App.Path & "\Pictures\SamatRotnest.jpg")
'loads and shows the picture of sam jumping from my file
End Sub

Private Sub cmdboat_Click()
picresults2.Picture = LoadPicture(App.Path & "\Pictures\LewinDocked.jpg")
'loads and shows the picture of the Lewin from my file
End Sub


Private Sub cmdclass_Click()
picresults2.Picture = LoadPicture(App.Path & "\Pictures\Homework.jpg")
'loads and displays the picture of us hard at work from my file
End Sub

Private Sub cmddance_Click()
picresults2.Picture = LoadPicture(App.Path & "\Pictures\SethandMeKP.jpg")
'loads and displays the picture of Seth and I before the dance from my file
End Sub

Private Sub cmdeat_Click()
picresults2.Picture = LoadPicture(App.Path & "\Pictures\TaraandIatColdRock.jpg")
'loads and shows the picture of Tara and I getting Fat from my file
End Sub

Private Sub cmdpub_Click()
picresults2.Picture = LoadPicture(App.Path & "\Pictures\Brashley2.jpg")
'loads and shows the picture of Brad and I at a pub from my file
End Sub

Private Sub cmdboom_Click()
picresults2.Picture = LoadPicture(App.Path & "\Pictures\AshleyChoppin2.jpg")
'loads and displays the picture of me making a boomerang from my file folder
End Sub

