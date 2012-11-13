VERSION 5.00
Begin VB.Form Questions_Phrases 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   7170
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   ScaleHeight     =   7170
   ScaleWidth      =   9480
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTip 
      BackColor       =   &H000000FF&
      Caption         =   """How much do you leave for a tip?"""
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4440
      Width           =   1935
   End
   Begin VB.CommandButton cmdExcuseMe 
      BackColor       =   &H000000FF&
      Caption         =   """Excuse me. Can we get the bill please?"""
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
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CommandButton cmdRestaurant 
      BackColor       =   &H000000FF&
      Caption         =   "Return to Restaurant Phrases"
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
      TabIndex        =   7
      Top             =   6120
      Width           =   1815
   End
   Begin VB.CommandButton cmdInDish 
      BackColor       =   &H000000FF&
      Caption         =   """What's in this dish here?"""
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
      TabIndex        =   6
      Top             =   5040
      Width           =   1815
   End
   Begin VB.CommandButton cmdNotEatAll 
      BackColor       =   &H000000FF&
      Caption         =   """I may not be able to eat it all. Can you wrap up any leftovers?"""
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1080
      Width           =   1815
   End
   Begin VB.CommandButton cmdLunch 
      BackColor       =   &H000000FF&
      Caption         =   """What's for lunch?"""
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
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2280
      Width           =   1815
   End
   Begin VB.CommandButton cmdWhatPlace 
      BackColor       =   &H000000FF&
      Caption         =   """Did you have a particular place in mind?"""
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
      TabIndex        =   3
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton cmdSoundsGood 
      BackColor       =   &H000000FF&
      Caption         =   """Sounds good. What are you in the mood for?"""
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
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3360
      Width           =   1815
   End
   Begin VB.CommandButton cmdBeenTo 
      BackColor       =   &H000000FF&
      Caption         =   """Have you been to the ___ ?"""
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2280
      Width           =   1935
   End
   Begin VB.CommandButton cmdRecommend 
      BackColor       =   &H000000FF&
      Caption         =   """Where do you recommend?"""
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
      TabIndex        =   0
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Label lblRestaurantQuestions 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Restaurant Questions"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   2760
      TabIndex        =   11
      Top             =   240
      Width           =   4335
   End
   Begin VB.Label lblIntro 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Click on the sentence for the Korean translation."
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   3120
      TabIndex        =   10
      Top             =   960
      Width           =   3495
   End
End
Attribute VB_Name = "Questions_Phrases"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Learning Korean 101
'Form Name: Questions_Phrases
'Amanda Phan and Natalie Hamilton
'Date: March 14
'Objective: The user has already clicked on the phrase category, Greetings, and the subcategory
    'Questions.  Now the user has the ability to click on the phrases shown in English,
    'and have the Korean phonetic translation.
'Comments: The form has buttons that the user can click on.  The user can click on the
    'desired phrase button and a messagebox will appear with the Korean phonetic translation.
    'If the user wishes to return to the previous form, they can click on the button labeled
    'return.  Once the user has clicked on the button, the present form will hide and
    'the desired previous form will appear.


Option Explicit

Private Sub cmdBeenTo_Click()
'This action will cause the button pushed to show the Korean translation in a messagebox.
MsgBox "_____ Sheek Dahng Ae Kah Baht Nee?", , "Translation"
End Sub

Private Sub cmdExcuseMe_Click()
'This action will cause the button pushed to show the Korean translation in a messagebox.
MsgBox "Sheel Lae Jee Mahn Kae Sahn Suh Johm Joo Sheel Lae Yo?", , "Translation"
End Sub

Private Sub cmdInDish_Click()
'This action will cause the button pushed to show the Korean translation in a messagebox.
MsgBox "Ee Yo Lee Ae Neun Moo Uh Shee Deu Luh Gah Nah Yo?", , "Translation"
End Sub

Private Sub cmdLunch_Click()
'This action will cause the button pushed to show the Korean translation in a messagebox.
MsgBox "Juhm Sheem Mae Nyoo Loh Muh Gah Eet Nah Yo?", , "Translation"
End Sub

Private Sub cmdNotEatAll_Click()
'This action will cause the button pushed to show the Korean translation in a messagebox.
MsgBox "Dah Moht Muh Geul Guht Gah Tah Yo. Nah Meun Guh Seun Ssah Joo Shee Gae Ssuh Yo?", , "Translation"
End Sub

Private Sub cmdRecommend_Click()
'This action will cause the button pushed to show the Korean translation in a messagebox.
MsgBox "Uh Dee Reul Choo Chuhn Ha Gaet Nee?", , "Translation"
End Sub

Private Sub cmdRestaurant_Click()
'This action will cause the present form to hide and the previous form to appear.
Questions_Phrases.Hide
Restaurants_Phrases.Show

End Sub

Private Sub cmdSoundsGood_Click()
'This action will cause the button pushed to show the Korean translation in a messagebox.

MsgBox "Keu Rae. Muh Muhk Goh Sheep Nee?", , "Translation"
End Sub

Private Sub cmdTip_Click()

'This action will cause the button pushed to show the Korean translation in a messagebox.
MsgBox "Tee Beu Loh Uhl Mah Noh Euh Myun Dwae Nah Yo?", , "Translation"
End Sub

Private Sub cmdWhatPlace_Click()
'This action will cause the button pushed to show the Korean translation in a messagebox.
MsgBox "Uh Dee Teuk Byul Hee Jung Hae Doon Sheek Dahng Ee Lah Doh Eet Nee?", , "Translation"
End Sub


