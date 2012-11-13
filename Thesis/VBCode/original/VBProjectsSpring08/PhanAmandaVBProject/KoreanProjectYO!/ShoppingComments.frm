VERSION 5.00
Begin VB.Form ShoppingComments 
   BackColor       =   &H000000FF&
   Caption         =   "Form1"
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7815
   LinkTopic       =   "Form1"
   ScaleHeight     =   6975
   ScaleWidth      =   7815
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      Height          =   4095
      Left            =   2520
      ScaleHeight     =   4035
      ScaleWidth      =   4635
      TabIndex        =   6
      Top             =   1800
      Width           =   4695
   End
   Begin VB.CommandButton cmdPerfect 
      BackColor       =   &H00FF0000&
      Caption         =   """This is perfect!"""
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton cmdILikey 
      BackColor       =   &H00FF0000&
      Caption         =   """I like this one here.  I think I'll take it."""
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton cmdShopAround 
      BackColor       =   &H00FF0000&
      Caption         =   """I think I'll shop around more."""
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton cmdIDontWant 
      BackColor       =   &H00FF0000&
      Caption         =   """I don't see anything I like."""
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton cmdNiceCheckOut 
      BackColor       =   &H00FF0000&
      Caption         =   """This is nice, check this out!"""
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00FF0000&
      Caption         =   "Return to Shopping"
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
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6120
      Width           =   1575
   End
   Begin VB.Label lblShopping 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Shopping"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2160
      TabIndex        =   8
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label lblIntro 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Please click on the follow buttons for the Korean translations!"
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
      Height          =   495
      Left            =   1920
      TabIndex        =   7
      Top             =   960
      Width           =   3855
   End
End
Attribute VB_Name = "ShoppingComments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Learning Korean 101
'Form Name: ShoppingComments
'Amanda Phan and Natalie Hamilton
'Date: March 10
'Objective: The user has already clicked on the phrase category, Shopping, and the subcategory
    'comments.  Here, the user can click on the English desired sentence and the Korean phonetic
    'translation will appear.
'Comments:  The form has buttons with English sentences on it. The user will have to click on the
    'the button and then in a messagebox, the Korean translation will appear. If the user would like
    'to return to the previous form, they will have to click on the button labeled return. When the
    'the button is clicked on, the present form will hide and the desired previous form
    'will appear.


Option Explicit

Private Sub cmdIDontWant_Click()
'This action will cause the button pushed to show the Korean translation in a messagebox.
MsgBox "Ee Joong Ae Suh Chae Gah Wohn Ha Neun Guhn Uhp Nae Yo", , "Translation"
End Sub

Private Sub cmdILikey_Click()
'This action will cause the button pushed to show the Korean translation in a messagebox.
MsgBox "Chuh Yuh Gee Ee Guh Mahm Ae Deu Nae Yo. Ee Guh Sah Goh Shee Puh Yo", , "Translation"
End Sub

Private Sub cmdNiceCheckOut_Click()
'This action will cause the button pushed to show the Korean translation in a messagebox.
MsgBox "Ee Guh Cham Kwaen Chan Nae Yo, Hahn Buhn Boh Sae Yo", , "Translation"
End Sub

Private Sub cmdPerfect_Click()
'This action will cause the button pushed to show the Korean translation in a messagebox.
MsgBox "Ee Guh Wahn Byuk Ha Nae Yo", , "Translation"
End Sub

Private Sub cmdReturn_Click()
'This action will cause the present form to hide and the previous form to appear.
ShoppingComments.Hide
Shopping.Show

End Sub

Private Sub cmdShopAround_Click()
'This action will cause the button pushed to show the Korean translation in a messagebox.
MsgBox "Johm Duh Dah Reun Kah Gae Deul Ae Guh Seul Ah Rah Bah Yah Gae Ssuh Yo", , "Translation"
End Sub

Private Sub Form_()
'This action will cause the picture to be loaded.
picResults.Picture = LoadPicture(App.Path & "\korean-dresses.jpg")
End Sub

Private Sub picResults_Click()
'This action will cause a picture to show on the form.
picResults.Picture = LoadPicture(App.Path & "\korean-dresses.jpg")
End Sub
