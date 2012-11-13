VERSION 5.00
Begin VB.Form BadMood 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   7650
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10035
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7650
   ScaleWidth      =   10035
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      Height          =   3495
      Left            =   3960
      ScaleHeight     =   3435
      ScaleWidth      =   2115
      TabIndex        =   13
      Top             =   2040
      Width           =   2175
   End
   Begin VB.CommandButton cmdWhyUpset 
      BackColor       =   &H000000FF&
      Caption         =   """Why are you upset?"""
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
      TabIndex        =   12
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton cmdVeryDisappointed 
      BackColor       =   &H000000FF&
      Caption         =   """I'm very disappointed."""
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
      TabIndex        =   11
      Top             =   4080
      Width           =   1695
   End
   Begin VB.CommandButton cmdOverboard 
      BackColor       =   &H000000FF&
      Caption         =   """I think you're going a little overboard."""
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5040
      Width           =   1695
   End
   Begin VB.CommandButton cmdQuitIt 
      BackColor       =   &H000000FF&
      Caption         =   """Quit it!"""
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
      TabIndex        =   9
      Top             =   3120
      Width           =   1695
   End
   Begin VB.CommandButton cmdImMad 
      BackColor       =   &H000000FF&
      Caption         =   """I'm really mad."""
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
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4080
      Width           =   1695
   End
   Begin VB.CommandButton cmdGetOut 
      BackColor       =   &H000000FF&
      Caption         =   """Get the heck out of here!"""
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
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3120
      Width           =   1695
   End
   Begin VB.CommandButton cmdJerk 
      BackColor       =   &H000000FF&
      Caption         =   """He's/She's such a jerk!"""
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
      TabIndex        =   6
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CommandButton cmdGonnaCry 
      BackColor       =   &H000000FF&
      Caption         =   """I think I'm going to cry."""
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
      TabIndex        =   5
      Top             =   6000
      Width           =   1695
   End
   Begin VB.CommandButton cmdMiserable 
      BackColor       =   &H000000FF&
      Caption         =   """I'm miserable."""
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
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton cmdDontLikeYOU 
      BackColor       =   &H000000FF&
      Caption         =   """I don't like you."""
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
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CommandButton cmdMadSaid 
      BackColor       =   &H000000FF&
      Caption         =   """I'm mad about what you said."""
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
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5040
      Width           =   1695
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H000000FF&
      Caption         =   "Return to Emotions"
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
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6840
      Width           =   1575
   End
   Begin VB.CommandButton cmdYouBadMood 
      BackColor       =   &H000000FF&
      Caption         =   """Because of you I'm in a bad mood."""
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6000
      Width           =   1695
   End
   Begin VB.Label lblIntro 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Please click on the sentence for the Korean Translation."
      BeginProperty Font 
         Name            =   "Franklin Gothic Demi"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   3000
      TabIndex        =   15
      Top             =   960
      Width           =   3975
   End
   Begin VB.Label lblBadMood 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Bad Mood"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   495
      Left            =   3120
      TabIndex        =   14
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "BadMood"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Learning Korean 101
'Form Name: BadMood
'Amanda Phan and Natalie Hamilton
'Date: March 7
'Objective: In this project, we have a phrase category called Emotions.
    'A subcategory we have is bad moods, and here are some English phrases
    'that will be translated into Korean phonetically from a messagebox.
'Comments: The form is used by the user clicking on the button that has the
    'desired English sentence.  Once the person has clicked on the button,
    'a messagebox will appear with the Korean phonetic translation. If the user
    'would like to return to the previous form/page, they just have to click on
    'the return button which will send them back to the previous form.


Option Explicit

Private Sub cmdDontLikeYOU_Click()
'This action will cause the button pushed to show the Korean translation in a messagebox.
MsgBox "Nuhl Ahn Joh Ah Hae Yo", , "Translation"
End Sub

Private Sub cmdGetOut_Click()
'This action will cause the button pushed to show the Korean translation in a messagebox.
MsgBox "Yuh Gee Suh Ssuhk Gguh Jyuh Buh Ryuh", , "Translation"
End Sub

Private Sub cmdGonnaCry_Click()
'This action will cause the button pushed to show the Korean translation in a messagebox.
MsgBox "Oo Luh Buh Leel Guht Gah Tah", , "Translation"
End Sub

Private Sub cmdImMad_Click()
'This action will cause the button pushed to show the Korean translation in a messagebox.
MsgBox "Jung Mal Mee Chyuh Buh Ree Gaet Nae", , "Translation"
End Sub

Private Sub cmdJerk_Click()
'This action will cause the button pushed to show the Korean translation in a messagebox.
MsgBox "Keu/Keu Nyuh Neun Hyung Pyun Uhb Neun Een Gah Nee Dah", , "Translation"
End Sub

Private Sub cmdMadSaid_Click()
'This action will cause the button pushed to show the Korean translation in a messagebox.
MsgBox "Nae Mal Ae Jung Mal Hwa Gah Nahn Dah", , "Translation"
End Sub

Private Sub cmdMiserable_Click()
'This action will cause the button pushed to show the Korean translation in a messagebox.
MsgBox "Bee Cham Han Kee Boon Ee Deul Uh", , "Translation"
End Sub

Private Sub cmdOverboard_Click()
'This action will cause the button pushed to show the Korean translation in a messagebox.
MsgBox "Nuh Johm Oh Buh Ha Neun Guht Gaht Dah", , "Translation"
End Sub

Private Sub cmdQuitIt_Click()
'This action will cause the button pushed to show the Korean translation in a messagebox.
MsgBox "Keu Mahn Dwoh", , "Translation"
End Sub

Private Sub cmdReturn_Click()
'This action will cause the present form to hide and the previous form to appear.
Emotions.Show
BadMood.Hide

End Sub

Private Sub cmdVeryDisappointed_Click()
'This action will cause the button pushed to show the Korean translation in a messagebox.
MsgBox "Nuh Moo Sheel Mahng Seu Ruh Woh", , "Translation"
End Sub

Private Sub cmdWhyUpset_Click()
'This action will cause the button pushed to show the Korean translation in a messagebox.
MsgBox "Nuh Wae Hwa Naht Nee?", , "Translation"
End Sub

Private Sub cmdYouBadMood_Click()
'This action will cause the button pushed to show the Korean translation in a messagebox.
MsgBox "Nuh Ddae Moon Ae Nahn Kee Boon Ee Nah Bbeu Dah", , "Translation"
End Sub

Private Sub Form_Load()
'This action will cause a picture to show on the form.
picResults.Picture = LoadPicture(App.Path & "\sad.jpg")

End Sub


Private Sub picResults_Click()
'This action will cause the picture to be loaded.
picResults.Picture = LoadPicture(App.Path & "\sad.jpg")
End Sub
