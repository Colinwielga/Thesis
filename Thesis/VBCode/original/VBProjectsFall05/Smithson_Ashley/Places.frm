VERSION 5.00
Begin VB.Form Places 
   BackColor       =   &H00000000&
   Caption         =   "Places"
   ClientHeight    =   8115
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10470
   LinkTopic       =   "Form4"
   ScaleHeight     =   8115
   ScaleWidth      =   10470
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdback 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Back to Home Page"
      BeginProperty Font 
         Name            =   "BatangChe"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6720
      Width           =   2175
   End
   Begin VB.PictureBox picresults 
      BackColor       =   &H00000000&
      Height          =   7335
      Left            =   240
      ScaleHeight     =   7275
      ScaleWidth      =   7035
      TabIndex        =   5
      Top             =   480
      Width           =   7095
   End
   Begin VB.CommandButton cmdperth 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Perth"
      BeginProperty Font 
         Name            =   "BatangChe"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5520
      Width           =   2175
   End
   Begin VB.CommandButton cmdfreo 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Fremantle"
      BeginProperty Font 
         Name            =   "BatangChe"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4320
      Width           =   2175
   End
   Begin VB.CommandButton cmdsing 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Singapore"
      BeginProperty Font 
         Name            =   "BatangChe"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3120
      Width           =   2175
   End
   Begin VB.CommandButton cmdbroome 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Broome"
      BeginProperty Font 
         Name            =   "BatangChe"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1920
      Width           =   2175
   End
   Begin VB.CommandButton cmdMel 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Melbourne"
      BeginProperty Font 
         Name            =   "BatangChe"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label Ashley 
      BackColor       =   &H00000000&
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
      TabIndex        =   7
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "Places"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Australia
'Form Name: Places
'Author: Ashley Smithson
'Date: October 31, 2005
'Purpose of form: to show pictures of just a few of the amazing sights you can see!
Option Explicit
Private Sub cmdback_Click()
Places.Hide
FinalProject2.Show
End Sub

Private Sub cmdbroome_Click()
picresults.Picture = LoadPicture(App.Path & "\Pictures\LionKing.jpg")
End Sub

Private Sub cmdfreo_Click()
picresults.Picture = LoadPicture(App.Path & "\Pictures\NightinFreo.jpg")
'find in folder and display picture
End Sub

Private Sub cmdMel_Click()
picresults.Picture = LoadPicture(App.Path & "\Pictures\Melbourne.jpg")
'ditto
End Sub

Private Sub cmdperth_Click()
picresults.Picture = LoadPicture(App.Path & "\Pictures\OverlookingPearthKP.jpg")
'ditto
End Sub

Private Sub cmdsing_Click()
picresults.Picture = LoadPicture(App.Path & "\Pictures\Singapore2.jpg")
'ditto
End Sub
