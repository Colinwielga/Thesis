VERSION 5.00
Begin VB.Form Naples 
   BackColor       =   &H0000C000&
   Caption         =   "Form6"
   ClientHeight    =   10575
   ClientLeft      =   1950
   ClientTop       =   450
   ClientWidth     =   11880
   LinkTopic       =   "Form6"
   ScaleHeight     =   10575
   ScaleWidth      =   11880
   Begin VB.CommandButton cmdreturn 
      Caption         =   "Look For Another Place to Adventure To?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8880
      TabIndex        =   6
      Top             =   2520
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      Height          =   5535
      Left            =   1200
      Picture         =   "Naples.frx":0000
      ScaleHeight     =   5475
      ScaleWidth      =   8955
      TabIndex        =   4
      Top             =   3840
      Width           =   9015
   End
   Begin VB.CommandButton cmdcalculate 
      Caption         =   "What Should I do Then"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7440
      TabIndex        =   3
      Top             =   1680
      Width           =   2355
   End
   Begin VB.TextBox txthours 
      Height          =   495
      Left            =   5040
      TabIndex        =   2
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label lblpompeii 
      BackColor       =   &H0000C000&
      Caption         =   "Have I left Enough Time To Visit Pompeii????"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2280
      TabIndex        =   5
      Top             =   2760
      Width           =   6015
   End
   Begin VB.Label lblvisit 
      BackColor       =   &H0000C000&
      Caption         =   "How many hours will you be in Naples?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1680
      TabIndex        =   1
      Top             =   1560
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C000&
      Caption         =   "What To Do In Naples?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
   End
End
Attribute VB_Name = "Naples"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Where to Travel in Italy
'Form Name: Naples
'Author: Sarah Dayton
'This form is to let the user know what their options are for Naples and finding out if they have enough time to go to Pompeii for how many hours they are willing to stay
Option Explicit

Private Sub cmdcalculate_Click()
Dim hours As Integer
hours = txthours
Select Case hours
    Case Is >= 10
        MsgBox ("Congratulations you have enought time to see both Pompeii and the National Archeological Musem!")
    Case 5 To 9
        MsgBox ("You have enough time to go to Pompeii, but not the National Archeological Museum.")
    Case 3, 4
        MsgBox ("Go to the National Archeological Museum in Naples!")
     Case Else
        MsgBox ("Guess you only get to walk around Naples.  Better check out some pizza places!")
 End Select
        
End Sub

Private Sub cmdreturn_Click()
OpeningPage.Show
Milan.Hide
Venice.Hide
Florence.Hide
Rome.Hide
Naples.Hide
SlideShowItaly.Hide
End Sub
