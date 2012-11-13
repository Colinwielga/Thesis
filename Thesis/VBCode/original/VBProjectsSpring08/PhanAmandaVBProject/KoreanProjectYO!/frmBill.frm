VERSION 5.00
Begin VB.Form Bill 
   BackColor       =   &H00008000&
   Caption         =   "Form1"
   ClientHeight    =   6315
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7980
   LinkTopic       =   "Form1"
   ScaleHeight     =   6315
   ScaleWidth      =   7980
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      Height          =   2535
      Left            =   1680
      ScaleHeight     =   2475
      ScaleWidth      =   5235
      TabIndex        =   4
      Top             =   720
      Width           =   5295
   End
   Begin VB.CommandButton cmdBigTip 
      BackColor       =   &H0000FF00&
      Caption         =   """Let's leave a big tip!"""
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
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton cmdMyTreat 
      BackColor       =   &H0000FF00&
      Caption         =   """It's my treat!"""
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
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton cmdMistakeBill 
      BackColor       =   &H0000FF00&
      Caption         =   """I think there's a mistake on this bill."""
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
      Left            =   840
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H0000FF00&
      Caption         =   "Return to Comment Phrases"
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
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5400
      Width           =   1575
   End
   Begin VB.Label lblIntro 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "Please click on the sentence for the Korean translaton."
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   1920
      TabIndex        =   6
      Top             =   3360
      Width           =   4215
   End
   Begin VB.Label lblBill 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "Bill"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "Bill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Learning Korean 101
'Form Name: Bill
'Amanda Phan and Natalie Hamilton
'Date: March 7
'Objective:  This project is about English to Korean translations.
    'Within these translations, there are English phrase categories dedicated to Restaurant
    'conversations.  In this subcategory of that, there are phrases dedicated
    'to the bill. Once clicked, the Korean phonetic translation appears.
'Comments:  This form is used by the user by clicking on the desired English sentence
    'button.  Once they click on it, a messagebox will appear with the Korean
    'phonetic translation. If they would like to return to the previous form/page, they
    'can click on the return page which will send the user back to the previous
    'form/page.

Option Explicit

Private Sub cmdBigTip_Click()
'This action will cause the button pushed to show the Korean translation in a messagebox.
MsgBox "Tee Beul Hoo Ha Gae Joo Jah", , "Translation"
End Sub

Private Sub cmdMistakeBill_Click()
'This action will cause the button pushed to show the Korean translation in a messagebox.
MsgBox "Kae Sahn Suh Ae Sheel Soo Gah Eet Neun Guht Gaht Seum Nee Dah", , "Translation"
End Sub

Private Sub cmdMyTreat_Click()
'This action will cause the button pushed to show the Korean translation in a messagebox.
MsgBox "Nae Gah 'This action will cause the button pushed to show the Korean translation in a messagebox.Nael Gae", , "Translation"
End Sub

Private Sub cmdReturn_Click()
'This action will cause the present form to hide and the previous form to appear.
Bill.Hide
Comment_Phrases.Show


End Sub

Private Sub Form_Load()
'This action will cause the picture to be loaded.

picResults.Picture = LoadPicture(App.Path & "\koreanmoney2.jpg")
End Sub

Private Sub picResults_Click()
'This action will cause a picture to show on the form.
picResults.Picture = LoadPicture(App.Path & "\koreanmoney2.jpg")
End Sub
