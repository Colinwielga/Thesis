VERSION 5.00
Begin VB.Form ShoppingQuestions 
   BackColor       =   &H000000FF&
   Caption         =   "Form1"
   ClientHeight    =   6900
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8670
   LinkTopic       =   "Form1"
   ScaleHeight     =   6900
   ScaleWidth      =   8670
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdHowMuch 
      BackColor       =   &H00FFFFFF&
      Caption         =   """How much is this?"""
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
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4320
      Width           =   1575
   End
   Begin VB.CommandButton cmdFittingRoom 
      BackColor       =   &H00FFFFFF&
      Caption         =   """Where is the fitting room?"""
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
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton cmdTryOn 
      BackColor       =   &H00FFFFFF&
      Caption         =   """Can I try this on?"""
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
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4320
      Width           =   1575
   End
   Begin VB.CommandButton cmdMySize 
      BackColor       =   &H00FFFFFF&
      Caption         =   """Do you have my size?"""
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
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton cmdPayWhere 
      BackColor       =   &H00FFFFFF&
      Caption         =   """Where can I pay for this?"""
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
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00FFFFFF&
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
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5520
      Width           =   1455
   End
   Begin VB.Label lblShoppingQuestions 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Shopping Questions"
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
      Left            =   2160
      TabIndex        =   7
      Top             =   480
      Width           =   4455
   End
   Begin VB.Label lblIntro 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Click on the sentence for the Korean translation."
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
      Left            =   2640
      TabIndex        =   6
      Top             =   1560
      Width           =   3615
   End
End
Attribute VB_Name = "ShoppingQuestions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Learning Korean 101
'Form Name: ShoppingQuestions
'Amanda Phan and Natalie Hamilton
'Date: March 12
'Objective: The user has already clicked on the phrase category, Shopping and the subcategory
    'Questions.  Here the user has English phrases to choose from and then when clicked,
    'the Korean phonetic translation will appear.
'Comments: The form has buttons that the user can click on.  Once the user has clicked
    'on the desired button, the Korean translation will appear in a messagebox. If the user would like
    'to return to the previous form, they will have to click on the button labeled return. When the
    'the button is clicked on, the present form will hide and the desired previous form
    'will appear.


Option Explicit
Private Sub cmdFittingRoom_Click()
'This action will cause the button pushed to show the Korean translation in a messagebox.
MsgBox "Oht Ee Buh Boh Neun Goh Shee Uh Dee Jyo?", , "Translation"
End Sub

Private Sub cmdHowMuch_Click()
'This action will cause the button pushed to show the Korean translation in a messagebox.
MsgBox "Ee Guht Uhl Mah Yae Yo?", , "Translation"
End Sub

Private Sub cmdMySize_Click()
'This action will cause the button pushed to show the Korean translation in a messagebox.
MsgBox "Chuh Han Tae Maht Neun Sah Ee Jeu Eet Nah Yo?", , "Translation"
End Sub

Private Sub cmdPayWhere_Click()
'This action will cause the button pushed to show the Korean translation in a messagebox.
MsgBox "Uh Dee Suh Kae Sahn Ha Myun Dwae Jyo?", , "Translation"
End Sub

Private Sub cmdReturn_Click()
'This action will cause the present form to hide and the previous form to appear.
Shopping.Show
ShoppingQuestions.Hide

End Sub

Private Sub cmdTryOn_Click()
'This action will cause the button pushed to show the Korean translation in a messagebox.
MsgBox "Ee Guh Ee Buh Bah Doh Dwae Nah Yo?", , "Translation"
End Sub


