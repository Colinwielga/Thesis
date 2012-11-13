VERSION 5.00
Begin VB.Form Emotions 
   BackColor       =   &H00FF0000&
   Caption         =   "Form1"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11070
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   11070
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      Height          =   5415
      Left            =   4560
      ScaleHeight     =   5355
      ScaleWidth      =   8115
      TabIndex        =   3
      Top             =   2400
      Width           =   8175
   End
   Begin VB.CommandButton cmdGoodMood 
      BackColor       =   &H000000FF&
      Caption         =   "Positive Emotions"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3240
      Width           =   1695
   End
   Begin VB.CommandButton cmdBadMood 
      BackColor       =   &H000000FF&
      Caption         =   "Negative Emotions"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3240
      Width           =   1815
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H000000FF&
      Caption         =   "Return to Phrases"
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
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Label lblEmotions 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Emotions"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   1920
      TabIndex        =   5
      Top             =   600
      Width           =   5175
   End
   Begin VB.Label lblIntro 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Please click on the category for more phrases!"
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
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   2160
      Width           =   4215
   End
End
Attribute VB_Name = "Emotions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Learning Korean 101
'Form Name: Emotions
'Amanda Phan and Natalie Hamilton
'Date: March 8
'Objective: This form is used for the category, Emotions.  Here, the user can choose
    'from 2 subcategories, Positive Emotions and Negative Emotions. The user can click
    'on the desired category button, and it will lead them to English phrases that
    'are related to the category.  There, they can get the Korean phonetic translation.
'Comments:  This form is used by the user have the choice of clicking on 2 buttons
    'that are based of two subcategories of the Emotion category.  Once clicked, these
    'buttons lead the user from that form to a new form that will have English phrases.
    'If the user would like to return to the previous form, the user can click on the
    'return button which will switch the form to the previous form.


Private Sub cmdBadMood_Click()
'This action will cause the present form to hide and the desired form to appear.
BadMood.Show
Emotions.Hide

End Sub

Private Sub cmdGoodMood_Click()
'This action will cause the present form to hide and the desired form to appear.
GoodMood.Show
Emotions.Hide

End Sub

Private Sub cmdReturn_Click()
'This action will cause the present form to hide and the previous form to appear.
Emotions.Hide
Phrases.Show

End Sub

Private Sub Form_Load()
'This action will cause the picture to be loaded.
picResults.Picture = LoadPicture(App.Path & "\happy.jpg")
End Sub

Private Sub picResults_Click()
'This action will cause a picture to show on the form.
picResults.Picture = LoadPicture(App.Path & "\happy.jpg")
End Sub
