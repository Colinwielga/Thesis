VERSION 5.00
Begin VB.Form Destination 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   6885
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8730
   BeginProperty Font 
      Name            =   "Book Antiqua"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6885
   ScaleWidth      =   8730
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdName 
      Caption         =   "Please enter you name :)"
      Height          =   615
      Left            =   3000
      TabIndex        =   7
      Top             =   1320
      Width           =   2295
   End
   Begin VB.PictureBox picResults 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8655
      Left            =   240
      ScaleHeight     =   8595
      ScaleWidth      =   6075
      TabIndex        =   4
      Top             =   2400
      Width           =   6135
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000000FF&
      Caption         =   "Quit"
      Height          =   855
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5760
      Width           =   2055
   End
   Begin VB.CommandButton cmdSwitch 
      BackColor       =   &H000000FF&
      Caption         =   "Return to Welcome Screen"
      Height          =   855
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7320
      Width           =   2055
   End
   Begin VB.CommandButton cmdDictionary 
      BackColor       =   &H000000FF&
      Caption         =   "Click here for the English/Korean Dictionary"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4440
      Width           =   2295
   End
   Begin VB.CommandButton cmdPhrases 
      BackColor       =   &H000000FF&
      Caption         =   "Click here for Korean Phrases"
      Height          =   855
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3240
      Width           =   2295
   End
   Begin VB.Label lblIntro 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "Please click on your destination!"
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   6600
      TabIndex        =   6
      Top             =   2040
      Width           =   3975
   End
   Begin VB.Label lblDestination 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "Your Korean knowledge begins here!"
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   1920
      TabIndex        =   5
      Top             =   480
      Width           =   4215
   End
End
Attribute VB_Name = "Destination"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Learning Korean 101
'Form Name: Destination
'Amanda Phan and Natalie Hamilton
'Date: March 7
'Objective:  This form allows the user to click on 2 categories that are dedicated to
    'having English to Korean phonetic translations.  Here, the user can click on the
    'phrases button which has already created English sentences dedicated to certain
    'categories.  The user can also click on the dictionary button which will lead them
    'a form that will allow the user to type in an English word, and have the Korean
    'phonetic translation appear.
'Comments:  The form is used by having the user click on either button.  Whatever button
    'they chose, it will lead the user to another form that is dedicated to either a dictionary
    'phrase categories.  If they would like to return to the start up screen, they just have
    'to click on the return button which will send them to the startup form.



Private Sub cmdDictionary_Click()
'This action will cause the present form to hide and the desired form to appear.
Destination.Hide
Dictionary.Show

End Sub



Private Sub cmdName_Click()
Dim Name As String
'This is where the person can input their name through an inputbox.
Name = InputBox("Please enter your name! :)")
'This is where the name shows up in a messagebox.
MsgBox ("Ahn young ha sae yo (hello!), " & Name & "! Welcome! Enjoy your stay!!")

End Sub

Private Sub cmdPhrases_Click()
'This action will cause the present form to hide and the desired form to appear.
Phrases.Show
Destination.Hide

End Sub

Private Sub cmdQuit_Click()
'This action will cause the person to quit the program.
End
End Sub

Private Sub cmdSwitch_Click()
'This action will cause the present form to hide and the previous form to appear.
Destination.Hide
Welcome.Show

End Sub

Private Sub Label1_Click()

End Sub

Private Sub Form_Load()
'This action will cause the picture to be loaded.

picResults.Picture = LoadPicture(App.Path & "\korean-map.jpg")
End Sub

Private Sub picResults_Click()
'This action will cause a picture to show on the form.
picResults.Picture = LoadPicture(App.Path & "\korean-map.jpg")
End Sub
