VERSION 5.00
Begin VB.Form Welcome 
   BackColor       =   &H000000FF&
   Caption         =   "Form1"
   ClientHeight    =   6150
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8550
   LinkTopic       =   "Form1"
   ScaleHeight     =   6150
   ScaleWidth      =   8550
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      Height          =   3375
      Left            =   1680
      ScaleHeight     =   3315
      ScaleWidth      =   5475
      TabIndex        =   2
      Top             =   1080
      Width           =   5535
   End
   Begin VB.CommandButton cmdEnter 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Click here to ENTER!!"
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
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4680
      Width           =   5895
   End
   Begin VB.Label lblWelcome 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Welcome to the Learning Korean Program!"
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
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   6975
   End
End
Attribute VB_Name = "Welcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Learning Korean 101
'Form Name: Welcome
'Amanda Phan and Natalie Hamilton
'Date: March 7
'Objective:  The main objective of this project was to have an English to Korean translation
    'program.  We decided on the basic level to have dictionary section, and then an already
    'created phrases section. Here then, a person could type in a word in a text box and
    'the Korean phonetic translation will appear.  Another way they could learn Korean
    'is the phrases we have already created. Here, the phrases are categorized and subcategorized.
    'The user can click on the buttons leading them to the actual phrases and then when the
    'phrase button is clicked, a messagebox will appear with the Korean phonetic translation.
    'This form is only used to introduce the program and for the user to enter the program.
'Comments:  The form has only one button to choose from, and this button allows the user
    'to enter the program.  Once the user has clicked on the button, the present form will
    'hide and the desired entrance from will appear.


Option Explicit
Private Sub cmdEnter_Click()
'This action will cause the present form to hide and the desired form to appear.
Welcome.Hide
Destination.Show

End Sub

Private Sub Form_Load()
'This action will cause a picture to show on the form.
picResults.Picture = LoadPicture(App.Path & "\South-Korea-flag.gif")
End Sub

Private Sub picResults_Click()
'This action will cause the picture to be loaded.
picResults.Picture = LoadPicture(App.Path & "\South-Korea-flag.gif")
End Sub
