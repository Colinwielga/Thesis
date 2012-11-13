VERSION 5.00
Begin VB.Form Shopping 
   BackColor       =   &H00FF0000&
   Caption         =   "Form1"
   ClientHeight    =   5715
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8070
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   8070
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      Height          =   4095
      Left            =   1680
      ScaleHeight     =   4035
      ScaleWidth      =   4755
      TabIndex        =   4
      Top             =   2280
      Width           =   4815
   End
   Begin VB.CommandButton cmdComments 
      BackColor       =   &H000000FF&
      Caption         =   "Comments/Answers"
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
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
      Width           =   1935
   End
   Begin VB.CommandButton cmdQuestions 
      BackColor       =   &H000000FF&
      Caption         =   "Questions"
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
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   960
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
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6960
      Width           =   1695
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
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   2160
      TabIndex        =   3
      Top             =   240
      Width           =   3495
   End
End
Attribute VB_Name = "Shopping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Learning Korean 101
'Form Name: Shopping
'Amanda Phan and Natalie Hamilton
'Date: March 9
'Objective: The user has already clicked on the phrase category, Shopping.  At this form,
    'they can choose on what type of phrases they would like see related to shopping.
    'They can choose either questions or comments.  Once they click on the desired button,
    'the desired subcategory form will appear.
'Comments: The form has 2 buttons that the user can choose to see more phrases.  These
    '2 buttons/phrases can be clicked on, and once they are, the present screen will
    'hide and the desired subcategory page(questions or comments) will appear. If the
    'user would like return to the previous form, they can click on the button labeled
    'return.  Once the user has clicked on the button, the present form will hide and
    'the desired previous form will appear.


Private Sub cmdComments_Click()
'This action will cause the present form to hide and the desired form to appear.
ShoppingComments.Show
Shopping.Hide

End Sub

Private Sub cmdQuestions_Click()
'This action will cause the present form to hide and the desired form to appear.
ShoppingQuestions.Show
Shopping.Hide

End Sub

Private Sub cmdReturn_Click()
'This action will cause the present form to hide and the previous form to appear.
Shopping.Hide
Phrases.Show

End Sub

Private Sub Form_Load()
'This action will cause the picture to be loaded.
picResults.Picture = LoadPicture(App.Path & "\korean-clothes.jpg")
End Sub

Private Sub picResults_Click()
'This action will cause a picture to show on the form.
picResults.Picture = LoadPicture(App.Path & "\korean-clothes.jpg")
End Sub
