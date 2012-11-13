VERSION 5.00
Begin VB.Form Restaurants_Phrases 
   BackColor       =   &H000000FF&
   ClientHeight    =   5910
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7830
   LinkTopic       =   "Form1"
   ScaleHeight     =   5910
   ScaleWidth      =   7830
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00FFFFFF&
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
      Height          =   735
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4080
      Width           =   1695
   End
   Begin VB.CommandButton cmdQuestions 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Click here for Questions Phrases"
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
      TabIndex        =   2
      Top             =   2280
      Width           =   1695
   End
   Begin VB.CommandButton cmdComments 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Click here for Comment Phrases"
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
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label lblIntro 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Click on the following button for the Korean translation of that sentence:"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   4695
   End
End
Attribute VB_Name = "Restaurants_Phrases"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Learning Korean 101
'Form Name: Restaurant_Phrases
'Amanda Phan and Natalie Hamilton
'Date: March 9
'Objective: The user has already clicked on the phrase category, Restaurants.  Now on this
    'form, the user can now choose on what type of restaurant phrases the would like.
    'The user can clicked on whether they would like to see questions or comments.
    'Once the user has clicked on the desired subcategory, they will be shown the
    'appropriate form.
'Comments:  This form has 2 buttons on it relating to restaurants.  Here, the user has
    'the choice of questions and comments, and then they can click on the desired
    'button.  Once they have clicked on the button, the present form will hide and
    'the desired form will appear.  If the user would like to return to the previous
    'form, they can click on the button labeled return, and the present form will hide
    'and the desired previous form will appear.



Private Sub cmdComments_Click()
'This action will cause the present form to hide and the desired form to appear.
Restaurants_Phrases.Hide
Comment_Phrases.Show

End Sub

Private Sub cmdQuestions_Click()
'This action will cause the present form to hide and the desired form to appear.
Restaurants_Phrases.Hide
Questions_Phrases.Show

End Sub

Private Sub cmdReturn_Click()
'This action will cause the present form to hide and the previous form to appear.
Restaurants_Phrases.Hide
Phrases.Show

End Sub


