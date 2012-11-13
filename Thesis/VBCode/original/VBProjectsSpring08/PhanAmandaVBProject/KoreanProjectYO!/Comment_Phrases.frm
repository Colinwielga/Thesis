VERSION 5.00
Begin VB.Form Comment_Phrases 
   BackColor       =   &H00800000&
   Caption         =   "Form1"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8085
   LinkTopic       =   "Form1"
   ScaleHeight     =   5865
   ScaleWidth      =   8085
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBill 
      BackColor       =   &H000000FF&
      Caption         =   "Bill/Tip"
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
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton cmdOrdering 
      BackColor       =   &H000000FF&
      Caption         =   "Ordering"
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
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton cmdFoodDescriptions 
      BackColor       =   &H000000FF&
      Caption         =   "Food Descriptions"
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
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4080
      Width           =   1695
   End
   Begin VB.CommandButton cmdRestaurant 
      BackColor       =   &H000000FF&
      Caption         =   "Return to Restaurant Phrases"
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
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5040
      Width           =   1695
   End
   Begin VB.CommandButton cmdFaveRest 
      BackColor       =   &H000000FF&
      Caption         =   """It's my favorite restaurant."""
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
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton cmdGetDinner 
      BackColor       =   &H000000FF&
      Caption         =   """Let's get dinner."""
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
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label lblDirections 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Please click on the category for more phrases."
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   1680
      TabIndex        =   8
      Top             =   3240
      Width           =   4455
   End
   Begin VB.Label lblIntro 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Please click on the sentence for the Korean translation.  "
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
      Left            =   2160
      TabIndex        =   7
      Top             =   840
      Width           =   3135
   End
   Begin VB.Label lblRestaurant 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Restaurant"
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
      Left            =   2160
      TabIndex        =   6
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "Comment_Phrases"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Learning Korean 101
'Form Name: Comment_Phrases
'Amanda Phan and Natalie Hamilton
'Date: March 7
'Objective: In this English to Korean translation project, there are categories that
    'the user can click on decicated to phrases used in that category. In this category
    'dedicated to restaurants, you can click on 2 English phrases commonly used and the
    'the Korean phonetic translation.  Another
    'aspect of this form is there are 3 more subcategories that the user can click on
    'focused on more specific conversations while in a restaurant.
'Comments: This form works by having the user have 2 choices on what they want to do.
    'If they would like to see the 2 English sentences translated into Korean, they can
    'click on the button and in messagebox, the Korean phonetic translation appears.
    'If they want to see phrases related to the categories listed, they can click
    'on the desired category button, which will send them to another form with the phrases.
    'If they would like to return back to the previous page, they just have to click
    'on the return button, and this will send the user back to the previous form/page.


Option Explicit


Private Sub cmdBill_Click()
'This action will cause the present form to hide and the desired form to appear.
Bill.Show
Comment_Phrases.Hide

End Sub

Private Sub cmdFaveRest_Click()
'This action will cause the button pushed to show the Korean translation in a messagebox.
MsgBox "Kuh Geen Nae Gah Kah Jahng Joh Ah Ha Neun Sheek Dahng Ee Yah", , "Translation"
End Sub

Private Sub cmdFoodDescriptions_Click()
'This action will cause the present form to hide and the desired form to appear.
FoodDescriptions.Show
Comment_Phrases.Hide

End Sub

Private Sub cmdGetDinner_Click()
'This action will cause the button pushed to show the Korean translation in a messagebox.
MsgBox "Juh Nyuk Sheek Sah Ha Ruh Kah Jah", , "Translation"
End Sub

Private Sub cmdOrdering_Click()
'This action will cause the present form to hide and the desired form to appear.
Ordering_Phrases.Show
Comment_Phrases.Hide

End Sub

Private Sub cmdRestaurant_Click()
'This action will cause the present form to hide and the previous form to appear.
Comment_Phrases.Hide
Restaurants_Phrases.Show

End Sub




Private Sub Form_Load()

End Sub
