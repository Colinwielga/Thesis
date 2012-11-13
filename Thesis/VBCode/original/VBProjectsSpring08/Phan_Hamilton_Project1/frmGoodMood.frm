VERSION 5.00
Begin VB.Form GoodMood 
   BackColor       =   &H00FF0000&
   Caption         =   "Form1"
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9570
   LinkTopic       =   "Form1"
   ScaleHeight     =   6975
   ScaleWidth      =   9570
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAwesome 
      BackColor       =   &H000000FF&
      Caption         =   """That's Awesome!"""
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
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4920
      Width           =   1575
   End
   Begin VB.CommandButton cmdLikeYou 
      BackColor       =   &H000000FF&
      Caption         =   """I like you."""
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
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4920
      Width           =   1575
   End
   Begin VB.CommandButton cmdLoveYou 
      BackColor       =   &H000000FF&
      Caption         =   """I love you."""
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
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton cmdWonderful 
      BackColor       =   &H000000FF&
      Caption         =   """Wonderful!"""
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
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3840
      Width           =   1455
   End
   Begin VB.CommandButton cmdCheerUp 
      BackColor       =   &H000000FF&
      Caption         =   """Cheer up!"""
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
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton cmdHappyYou 
      BackColor       =   &H000000FF&
      Caption         =   """""I'm really happy for you."""
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
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton cmdHappyNoSpeak 
      BackColor       =   &H000000FF&
      Caption         =   """I'm so happy, I do not know what to say!"""
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
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton cmdSoHappy 
      BackColor       =   &H000000FF&
      Caption         =   """I'm so happy!"""
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
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H000000FF&
      Caption         =   "Return to Emotions"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6000
      Width           =   1455
   End
   Begin VB.Label lblGoodMood 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Good Mood"
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
      Left            =   2160
      TabIndex        =   10
      Top             =   240
      Width           =   5175
   End
   Begin VB.Label lblIntro 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
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
      Height          =   375
      Left            =   3000
      TabIndex        =   9
      Top             =   1080
      Width           =   3855
   End
End
Attribute VB_Name = "GoodMood"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Learning Korean 101
'Form Name: GoodMood
'Amanda Phan and Natalie Hamilton
'Date: March 12
'Objective: In this form, the user has already clicked on the Emotions category, and now the
    'positive emotions subcategory.  Here, the user can click on the phrases' buttons that
    'are related to the positve emotions, and the
    'Korean phonetic translation will appear.
'Comments: In this form, the user will have the choice to click on the desired button
    'that features an English phrase.  Once the user has clicked on the button, a
    'messagebox will appear with the Korean Phonetic translation.  If the user would
    'like to return to the previous form, the user can click on the button labeled
    'return.  Here, once the button is clicked, the present form will hide and the
    'desired previous form will appear.


Option Explicit


Private Sub cmdAwesome_Click()
'This action will cause the button pushed to show the Korean translation in a messagebox.
MsgBox "Kwaeng Jahng Han Dae!", , "Translation"
End Sub

Private Sub cmdCheerUp_Click()
'This action will cause the button pushed to show the Korean translation in a messagebox.
MsgBox "Kee Oon Nae", , "Translation"
End Sub



Private Sub cmdHappyNoSpeak_Click()
'This action will cause the button pushed to show the Korean translation in a messagebox.
MsgBox "Nuh Moo Kee Bbuh Suh Moo Seun Mah Leul Hae Yah Hal Jee Moh Reu Gae Ssuh", , "Translation"
End Sub

Private Sub cmdHappyYou_Click()
'This action will cause the button pushed to show the Korean translation in a messagebox.
MsgBox "Nae Gah Jal Dwae Suh Nah Doh Kee Bbuh!", , "Translation"
End Sub

Private Sub cmdLikeYou_Click()
'This action will cause the button pushed to show the Korean translation in a messagebox.
MsgBox "Nahn Nuhl Joh Ah Hae Yo", , "Translation"
End Sub

Private Sub cmdLoveYou_Click()
'This action will cause the button pushed to show the Korean translation in a messagebox.
MsgBox "Sah Rahng Hae Yo", , "Translation"
End Sub

Private Sub cmdReturn_Click()

'This action will cause the present form to hide and the previous form to appear.
GoodMood.Hide
Emotions.Show

End Sub

Private Sub cmdSoHappy_Click()
'This action will cause the button pushed to show the Korean translation in a messagebox.
MsgBox "Nahn Mae Oo Kee Bbuh", , "Translation"
End Sub

Private Sub cmdWonderful_Click()
'This action will cause the button pushed to show the Korean translation in a messagebox.
MsgBox "Muht Jee Nae Yo!", , "Translation"
End Sub

