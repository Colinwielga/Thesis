VERSION 5.00
Begin VB.Form Greetings 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   5385
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8475
   LinkTopic       =   "Form1"
   ScaleHeight     =   5385
   ScaleWidth      =   8475
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      Height          =   2655
      Left            =   4080
      ScaleHeight     =   2595
      ScaleWidth      =   2595
      TabIndex        =   3
      Top             =   1200
      Width           =   2655
   End
   Begin VB.CommandButton cmdGreetingQuestions 
      BackColor       =   &H000000FF&
      Caption         =   "Click Here to Go to Greeting Questions"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2880
      Width           =   1815
   End
   Begin VB.CommandButton cmdGreetingComments 
      BackColor       =   &H000000FF&
      Caption         =   "Click Here to Go to Greeting Comments"
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
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H000000FF&
      Caption         =   "Return to Phrases"
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
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Label lblGreetings 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Greetings"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   2040
      TabIndex        =   5
      Top             =   240
      Width           =   4095
   End
   Begin VB.Label lblIntro 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Click on the category for more Korean phrases."
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   3495
   End
End
Attribute VB_Name = "Greetings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Learning Korean 101
'Form Name: Greetings
'Amanda Phan and Natalie Hamilton
'Date: March 9
'Objective:  In this form, the user has already clicked on the phrase category, Greetings.
    'The user now has the choice on what subcategory they would like, Greeting Comments
    'or Greetings Questions.  Once the user has clicked on the desired subcategory
    'button, they will go to the desired form.
'Comments:  In this form, the user has the choice of what subcategory they would like
    'to see. There are 2 subcategory buttons they can choose from, and once the button
    'is clicked, the present form is hidden, and the desired form appears.  If the
    'user would like to return to the previous form, they would click on the button
    'labeled return.  Once the user has clicked on that, the present form will hide
    'and the desired previous form will appear.




Private Sub cmdGreetingComments_Click()
'This action will cause the present form to hide and the desired form to appear.
Greetings.Hide
GreetingComments.Show

End Sub

Private Sub cmdGreetingQuestions_Click()
'This action will cause the present form to hide and the desired form to appear.
Greetings.Hide
GreetingQuestions.Show

End Sub

Private Sub cmdReturn_Click()
'This action will cause the present form to hide and the previous form to appear.
Greetings.Hide
Phrases.Show

End Sub

Private Sub Form_Load()
'This action will cause the picture to be loaded.
picResults.Picture = LoadPicture(App.Path & "\greeting_1.jpg")
End Sub

Private Sub picResults_Click()
'This action will cause a picture to show on the form.
picResults.Picture = LoadPicture(App.Path & "\greeting_1.jpg")
End Sub
