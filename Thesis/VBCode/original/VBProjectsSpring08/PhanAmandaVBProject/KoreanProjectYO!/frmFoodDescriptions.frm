VERSION 5.00
Begin VB.Form FoodDescriptions 
   BackColor       =   &H000000FF&
   Caption         =   "Form1"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9405
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   9405
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      Height          =   3975
      Left            =   2880
      ScaleHeight     =   3915
      ScaleWidth      =   3555
      TabIndex        =   6
      Top             =   1560
      Width           =   3615
   End
   Begin VB.CommandButton cmdStarving 
      BackColor       =   &H00FF0000&
      Caption         =   """I'm starving!"""
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2400
      Width           =   1815
   End
   Begin VB.CommandButton cmdFoodBad 
      BackColor       =   &H00FF0000&
      Caption         =   """The food was bad."""
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4200
      Width           =   1935
   End
   Begin VB.CommandButton cmdStuffed 
      BackColor       =   &H00FF0000&
      Caption         =   """I'm stuffed."""
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1800
      Width           =   1935
   End
   Begin VB.CommandButton cmdFoodGood 
      BackColor       =   &H00FF0000&
      Caption         =   """The food was good."""
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3000
      Width           =   1935
   End
   Begin VB.CommandButton cmdNotHungry 
      BackColor       =   &H00FF0000&
      Caption         =   """I'm really not that hungry."""
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3720
      Width           =   1815
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00FF0000&
      Caption         =   "Return to Comment Phrases"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5760
      Width           =   1695
   End
   Begin VB.Label lblFoodDescriptions 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Food Descriptions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   2400
      TabIndex        =   8
      Top             =   240
      Width           =   4695
   End
   Begin VB.Label lblIntro 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Please click on the sentence for the Korean translation."
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   2640
      TabIndex        =   7
      Top             =   960
      Width           =   4095
   End
End
Attribute VB_Name = "FoodDescriptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Learning Korean 101
'Form Name: FoodDescriptions
'Amanda Phan and Natalie Hamilton
'Date: March 10
'Objective: In this form, the user has already clicked on the phrase category, Restaurants.
    'The user has already clicked on the subcategory, Food Descriptions, which now the
    'user can click on the desired English sentence, and the Korean phonetic translation
    'appears.
'Comments:  This form is used by having the form have buttons the user can click on.
    'These buttons have English phrases on them, and when the user has clicked on the
    'button, the Korean phonetic translation will appear in a messagebox. If the user
    'would like to return to the previous page, then they would click on the button labeled
    'return.  When the user clicks on that button, the present page hides and the desired
    'previous form appears.


Option Explicit

Private Sub cmdFoodBad_Click()
'This action will cause the button pushed to show the Korean translation in a messagebox.
MsgBox "Mah Shee Uhb Ssuh Ssuh Yo", , "Translation"
End Sub

Private Sub cmdFoodGood_Click()
'This action will cause the button pushed to show the Korean translation in a messagebox.
MsgBox "Mah Shee Ssuh Ssuh Yo", , "Translation"
End Sub

Private Sub cmdNotHungry_Click()
'This action will cause the button pushed to show the Korean translation in a messagebox.
MsgBox "Nahn Keu Jung Doh Roh Pae Goh Peu Jee Neun Ahn Ah", , "Translation"
End Sub

Private Sub cmdReturn_Click()
'This action will cause the present form to hide and the previous form to appear.
FoodDescriptions.Hide
Comment_Phrases.Show

End Sub

Private Sub cmdStarving_Click()
'This action will cause the button pushed to show the Korean translation in a messagebox.
MsgBox "Nah Pae Goh Pah Joo Keul Jee Gyung Ee Yah", , "Translation"
End Sub

Private Sub cmdStuffed_Click()
'This action will cause the button pushed to show the Korean translation in a messagebox.
MsgBox "Pae Gah Kkwak Cha Ssuh", , "Translation"
End Sub


Private Sub Form_Load()
'This action will cause the picture to be loaded.
picResults.Picture = LoadPicture(App.Path & "\korean_bbq_puchong_3.jpg")
End Sub

Private Sub picResults_Click()
'This action will cause a picture to show on the form.
picResults.Picture = LoadPicture(App.Path & "\korean_bbq_puchong_3.jpg")
End Sub
