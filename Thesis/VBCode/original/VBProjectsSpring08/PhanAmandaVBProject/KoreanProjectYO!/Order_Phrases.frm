VERSION 5.00
Begin VB.Form Ordering_Phrases 
   BackColor       =   &H00FF0000&
   Caption         =   "Form1"
   ClientHeight    =   7215
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7215
   ScaleWidth      =   9000
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSameOrder 
      BackColor       =   &H000000FF&
      Caption         =   """I'll have the same."""
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
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4080
      Width           =   1815
   End
   Begin VB.CommandButton cmdPickyEater 
      BackColor       =   &H000000FF&
      Caption         =   """I'm kind of a picky eater."""
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
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3600
      Width           =   1815
   End
   Begin VB.CommandButton cmdNeverBeen 
      BackColor       =   &H000000FF&
      Caption         =   """I've never been here before."""
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
      TabIndex        =   3
      Top             =   2640
      Width           =   1815
   End
   Begin VB.CommandButton cmdHighRecommend 
      BackColor       =   &H000000FF&
      Caption         =   """I highly recommend it."""
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
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4800
      Width           =   1815
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H000000FF&
      Caption         =   "Return to Comments"
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
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6000
      Width           =   1695
   End
   Begin VB.CommandButton cmdConfused 
      BackColor       =   &H000000FF&
      Caption         =   """I haven't made up my mind yet.  Everything looks so good."""
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
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label lblOrdering 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Ordering"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2040
      TabIndex        =   7
      Top             =   360
      Width           =   4575
   End
   Begin VB.Label lblIntro 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Click on the sentence for the Korean translation."
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   2160
      TabIndex        =   6
      Top             =   1320
      Width           =   4575
   End
End
Attribute VB_Name = "Ordering_Phrases"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Learning Korean 101
'Form Name: Ordering_Phrases
'Amanda Phan and Natalie Hamilton
'Date: March 11
'Objective: The user has already clicked on the phrase category, Restauarant
    'and the subcategory, Ordering.  Now the user has the choices
    'of what phrases they would like to see for the Korean phonetic translation.
    'The user would click on the desired button and the translation will apppear.
'Comments:  The form is allows the user to click on the buttons on the form.
    'Once the user has clicked on the button, the Korean translation will
    'appear through a messagebox.  If the user would like to return to the previous
    'form they can click on the button labeled return. Once they have clicked on the
    'button, the present form will hide and the desired previous form will appear.



Option Explicit

Private Sub cmdConfused_Click()
'This action will cause the button pushed to show the Korean translation in a messagebox.
MsgBox "Ah Jeek Moht Gohl Lah Ssuh Yo. Dah Mah Shee Suh Boh Ee Neun Dae Yo", , "Translation"
End Sub

Private Sub cmdHighRecommend_Click()
'This action will cause the button pushed to show the Korean translation in a messagebox.
MsgBox "Chuhk Geuk Choo Chuhn Hahm Nee Dah", , "Translation"
End Sub

Private Sub cmdNeverBeen_Click()
'This action will cause the button pushed to show the Korean translation in a messagebox.
MsgBox "Chuhn Yuh Gee Juhn Ae Wah Bohn Juhk Ee Uhb Suh Yo", , "Translation"
End Sub

Private Sub cmdPickyEater_Click()
'This action will cause the button pushed to show the Korean translation in a messagebox.
MsgBox "Nahn Eum Shee Geul Dae Chae Loh Kah Lee Neun Pyun Ee Yah", , "Translation"
End Sub

Private Sub cmdReturn_Click()
'This action will cause the present form to hide and the previous form to appear.
Ordering_Phrases.Hide
Comment_Phrases.Show

End Sub

Private Sub cmdSameOrder_Click()
'This action will cause the button pushed to show the Korean translation in a messagebox.
MsgBox "Chuh Doh Gah Teun Guhl Loh Joo Sae Yo", , "Translation"
End Sub


Private Sub Form_Load()

End Sub
