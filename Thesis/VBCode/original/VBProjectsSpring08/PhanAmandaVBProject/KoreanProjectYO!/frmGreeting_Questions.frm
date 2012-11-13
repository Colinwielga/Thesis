VERSION 5.00
Begin VB.Form GreetingQuestions 
   BackColor       =   &H00FF0000&
   Caption         =   "Form1"
   ClientHeight    =   7440
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9690
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   9690
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFine 
      BackColor       =   &H000000FF&
      Caption         =   """Fine, thanks. And you?"""
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
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5160
      Width           =   1815
   End
   Begin VB.CommandButton cmdWhatsWrong 
      BackColor       =   &H000000FF&
      Caption         =   """What's wrong?"""
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
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4200
      Width           =   1815
   End
   Begin VB.CommandButton cmdDoingHere 
      BackColor       =   &H000000FF&
      Caption         =   """What are you doing here?"""
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
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3240
      Width           =   1695
   End
   Begin VB.CommandButton cmdExcuseMe 
      BackColor       =   &H000000FF&
      Caption         =   """Excuse me, what's your name?"""
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
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3240
      Width           =   1815
   End
   Begin VB.CommandButton cmdEverything_Ok 
      BackColor       =   &H000000FF&
      Caption         =   """You don't look so good. Is everything OK?"""
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
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5160
      Width           =   1815
   End
   Begin VB.CommandButton cmdHowIsDay 
      BackColor       =   &H000000FF&
      Caption         =   """How was your day?"""
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
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5160
      Width           =   1695
   End
   Begin VB.CommandButton cmdWeekend 
      BackColor       =   &H000000FF&
      Caption         =   """How was your weekend?"""
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
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4200
      Width           =   1815
   End
   Begin VB.CommandButton cmdLongTimeNoSee 
      BackColor       =   &H000000FF&
      Caption         =   """Long time no see! How have you been?"""
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
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4200
      Width           =   1815
   End
   Begin VB.CommandButton cmdHowLong 
      BackColor       =   &H000000FF&
      Caption         =   """How long has it been  since we last saw each other?"""
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
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5160
      Width           =   1815
   End
   Begin VB.CommandButton cmdHello 
      BackColor       =   &H000000FF&
      Caption         =   """Hello, how are you?"""
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
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4200
      Width           =   1695
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Return to Greetings"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6120
      Width           =   2055
   End
   Begin VB.Label lblGreetingQuestions 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Greeting Questions"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2760
      TabIndex        =   12
      Top             =   360
      Width           =   4215
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
      Left            =   2880
      TabIndex        =   11
      Top             =   1800
      Width           =   3855
   End
End
Attribute VB_Name = "GreetingQuestions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Learning Korean 101
'Form Name: GreetingQuestions
'Amanda Phan and Natalie Hamilton
'Date: March 13
'Objective: In this form, the user has already clicked on the Greetings phrases category,
    'and the Greetings Questions subcategory. Now the user has the choice of clicking
    'on the desired English phrase button.  Once the button is clicked, the Korean
    'phonetic translation will appear.
'Comments:  In this form, the user has buttons to choice from that have English phrases
    'on them.  Here, the user can click on these buttons, in the Korean phonetic translation
    'will appear in a messagebox. The


Option Explicit

Private Sub cmdDoingHere_Click()
'This action will cause the button pushed to show the Korean translation in a messagebox.
MsgBox "Nuh Yuh Gee Suh Muh Ha Nee?", , "Translation"
End Sub

Private Sub cmdEverything_Ok_Click()
'This action will cause the button pushed to show the Korean translation in a messagebox.
MsgBox "Byul Loh Joh Ah Boh Ee Jeel Ahn Nae Yo. Kwaen Chan Euh Sae Yo?", , "Translation"
End Sub

Private Sub cmdExcuseMe_Click()
'This action will cause the button pushed to show the Korean translation in a messagebox.
MsgBox "Sheel Lae Hahm Nee Dah, Ee Reum Ee Moo Uht Sheem Nee Kka?", , "Translation"
End Sub

Private Sub cmdFine_Click()
'This action will cause the button pushed to show the Korean translation in a messagebox.
MsgBox "Jal Jee Nae Jyo. Daek Eun Yo?", , "Translation"
End Sub

Private Sub cmdHello_Click()
'This action will cause the button pushed to show the Korean translation in a messagebox.
MsgBox "Ahn Young Ha Sae Yo", , "Translation"
End Sub

Private Sub cmdHowIsDay_Click()
'This action will cause the button pushed to show the Korean translation in a messagebox.
MsgBox "Oh Neul Ha Roo Uh Dae Suh Yo?", , "Translation"
End Sub

Private Sub cmdHowLong_Click()
'This action will cause the button pushed to show the Korean translation in a messagebox.
MsgBox "Juhn Ae Mahn Nah Go Nah Suh Ee Gae Uhl Mah Mahn Ee Ae Yo?", , "Translation"
End Sub

Private Sub cmdLongTimeNoSee_Click()
'This action will cause the button pushed to show the Korean translation in a messagebox.
MsgBox "Oh Raen Mahn Ee Yah! Uh Duh Kae Jee Nae Suh?", , "Translation"
End Sub

Private Sub cmdReturn_Click()
'This action will cause the present form to hide and the previous form to appear.
Greetings.Show
GreetingQuestions.Hide
End Sub

Private Sub cmdWeekend_Click()
'This action will cause the button pushed to show the Korean translation in a messagebox.
MsgBox "Joo Mal Jal Jee Nae Ssuh Yo?", , "Translation"
End Sub

Private Sub cmdWhatsWrong_Click()
'This action will cause the button pushed to show the Korean translation in a messagebox.
MsgBox "Moo Seun Eel Ee Suh Yo?", , "Translation"
End Sub

Private Sub Form_Load()

End Sub
