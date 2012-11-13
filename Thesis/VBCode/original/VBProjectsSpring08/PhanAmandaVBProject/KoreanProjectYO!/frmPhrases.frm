VERSION 5.00
Begin VB.Form Phrases 
   BackColor       =   &H00FF0000&
   Caption         =   "Form1"
   ClientHeight    =   6600
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9270
   LinkTopic       =   "Form1"
   ScaleHeight     =   6600
   ScaleWidth      =   9270
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H000000FF&
      Caption         =   "Return to Destination"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5400
      Width           =   1695
   End
   Begin VB.CommandButton cmdShopping 
      BackColor       =   &H000000FF&
      Caption         =   "Shopping"
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
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton cmdEmotions 
      BackColor       =   &H000000FF&
      Caption         =   "Emotions"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3600
      Width           =   1575
   End
   Begin VB.CommandButton cmdGreetings 
      BackColor       =   &H000000FF&
      Caption         =   "Greetings"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton cmdRestaurants 
      BackColor       =   &H000000FF&
      Caption         =   "Restaurants"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label lblIntro 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Click on the following buttons for Korean phrases in those situations:"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   2040
      TabIndex        =   2
      Top             =   600
      Width           =   5295
   End
End
Attribute VB_Name = "Phrases"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Learning Korean 101
'Form Name: Phrases
'Amanda Phan and Natalie Hamilton
'Date: March 13
'Objective: The user will already have clicked on the choice to see English to Korean
    'phrases.  Here, the user will now have the choices of what categories
    'of phrases they would like to see.  They can click on the four categories: Restaurant,
    'Greetings, Shopping, and Emotions.  Once they click on the desired category button,
    'they will go to the desired category form.
'Comments:  The form has 4 buttons that the user can click on.  These 4 buttons represent
    'the 4 categories they can click on. Once the user has clicked on the 4 buttons, the
    'present form will hide and the desired category form will appear.  If the user
    'would like to return to the startup page, the user can click on the button labeled
    'return.  Once the button is clicked, the present form will hide and the desired
    'previous form will appear.


Private Sub cmdEmotions_Click()
'This action will cause the present form to hide and the desired form to appear.
Emotions.Show
Phrases.Hide

End Sub

Private Sub cmdGreetings_Click()
'This action will cause the present form to hide and the desired form to appear.
Greetings.Show
Phrases.Hide

End Sub

Private Sub cmdRestaurants_Click()
'This action will cause the present form to hide and the desired form to appear.
Restaurants_Phrases.Show
Phrases.Hide

End Sub

Private Sub cmdReturn_Click()
'This action will cause the present form to hide and the previous form to appear.
Destination.Show
Phrases.Hide

End Sub

Private Sub cmdShopping_Click()
'This action will cause the present form to hide and the desired form to appear.
Phrases.Hide
Shopping.Show

End Sub

Private Sub Form_Load()

End Sub
