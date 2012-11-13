VERSION 5.00
Begin VB.Form GreetingComments 
   BackColor       =   &H000000FF&
   Caption         =   "Form1"
   ClientHeight    =   7680
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10035
   LinkTopic       =   "Form1"
   ScaleHeight     =   7680
   ScaleWidth      =   10035
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDontSeeMuch 
      BackColor       =   &H00FF0000&
      Caption         =   """I don't see you much these days."""
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
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3840
      Width           =   1695
   End
   Begin VB.CommandButton cmdSmallWorld 
      BackColor       =   &H00FF0000&
      Caption         =   """It's a small world."""
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
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2760
      Width           =   1695
   End
   Begin VB.CommandButton cmdGood 
      BackColor       =   &H00FF0000&
      Caption         =   """I'm good."""
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
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton cmdBusy 
      BackColor       =   &H00FF0000&
      Caption         =   """I'm kind of busy."""
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
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5160
      Width           =   1695
   End
   Begin VB.CommandButton cmdHeardALot 
      BackColor       =   &H00FF0000&
      Caption         =   """I've heard a lot about you."""
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
      TabIndex        =   5
      Top             =   5160
      Width           =   1935
   End
   Begin VB.CommandButton cmdGreatToSeeAgain 
      BackColor       =   &H00FF0000&
      Caption         =   """It's great to see you again. We have to get together more often."""
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5160
      Width           =   1935
   End
   Begin VB.CommandButton cmdCoincidence 
      BackColor       =   &H00FF0000&
      Caption         =   """What a coincidence!"""
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
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3840
      Width           =   1695
   End
   Begin VB.CommandButton cmdNothing 
      BackColor       =   &H00FF0000&
      Caption         =   """Nothing, not much."""
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
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2760
      Width           =   1695
   End
   Begin VB.CommandButton cmdFirstMeeting 
      BackColor       =   &H00FF0000&
      Caption         =   """Nice to meet you."""
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
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00FF0000&
      Caption         =   "Return to Greetings"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6360
      Width           =   2055
   End
   Begin VB.Label lblGreetings 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2640
      TabIndex        =   11
      Top             =   360
      Width           =   4815
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3000
      TabIndex        =   10
      Top             =   1200
      Width           =   4215
   End
End
Attribute VB_Name = "GreetingComments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Learning Korean 101
'Form Name: GreetingsComments
'Amanda Phan and Natalie Hamilton
'Date: March 10
'Objective: In this form, the user has already clicked on the Greetings category and
    'the Greeting Comments subcategory.  Now, the user has buttons that have English
    'phrases that are related to greetings but are not questions.  The user can click
    'on the desired phrase button and the Korean phonetic translation will appear.
'Comments:  In this form, there are many buttons that have English phrases on it.  The
    'must click on the button to get the Korean translation.  Here, once the button
    'has been clicked, then the Korean phonetic translation will appear in a messagebox.
    'If the user would like to return to the previous form, the user can click on the
    'button labeled return.  Here, the present form will hide and the desired previous
    'form will appear.


Option Explicit

Private Sub cmdBusy_Click()
'This action will cause the button pushed to show the Korean translation in a messagebox.
MsgBox "Johm Bah Bbah Yo", , "Translation"
End Sub

Private Sub cmdCoincidence_Click()
'This action will cause the button pushed to show the Korean translation in a messagebox.
MsgBox "Ee Ruhn Oo Yuhn Ee Eet Dah Nee", , "Translation"
End Sub

Private Sub cmdDontSeeMuch_Click()
'This action will cause the button pushed to show the Korean translation in a messagebox.
MsgBox "Yo Jeum Tong Mot Mahn Naht Nae Yo", , "Translation"
End Sub

Private Sub cmdFirstMeeting_Click()
'This action will cause the button pushed to show the Korean translation in a messagebox.
MsgBox "Mahn Nah Suh Pahn Gahp Seum Nee Dah", , "Translation"
End Sub

Private Sub cmdGood_Click()
'This action will cause the button pushed to show the Korean translation in a messagebox.
MsgBox "Jal Jee Nae Yo", , "Translation"
End Sub

Private Sub cmdGreatToSeeAgain_Click()
'This action will cause the button pushed to show the Korean translation in a messagebox.
MsgBox "Dah Shee Mahn Nah Nee Ggah Jung Mal Pahn Gahp Dah. Ee Jae Johm Duh Jah Joo Mahn Nah Sseu Myun Joh Kae Ssuh", , "Translation"
End Sub

Private Sub cmdHeardALot_Click()
'This action will cause the button pushed to show the Korean translation in a messagebox.
MsgBox "Mal Sseum Mahn Ee Deul Uht Seum Nee Dah", , "Translation"
End Sub

Private Sub cmdNothing_Click()
'This action will cause the button pushed to show the Korean translation in a messagebox.
MsgBox "Keu Ruhk Juh Ruhk", , "Translation"
End Sub

Private Sub cmdReturn_Click()
'This action will cause the present form to hide and the previous form to appear.
GreetingComments.Hide
Greetings.Show
End Sub

Private Sub cmdSmallWorld_Click()
'This action will cause the button pushed to show the Korean translation in a messagebox.
MsgBox "Sae Sahng Cham Chop Dah", , "Translation"
End Sub


Private Sub Form_Load()

End Sub
