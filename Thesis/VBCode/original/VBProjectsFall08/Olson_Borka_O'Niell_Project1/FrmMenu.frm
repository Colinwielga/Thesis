VERSION 5.00
Begin VB.Form FrmForm1 
   BackColor       =   &H00FFFF00&
   Caption         =   "Form1"
   ClientHeight    =   8730
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14010
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "FrmMenu.frx":0000
   ScaleHeight     =   8730
   ScaleWidth      =   14010
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Cmdsearch 
      Caption         =   "Search"
      Height          =   975
      Left            =   1800
      TabIndex        =   9
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Display"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9960
      TabIndex        =   8
      Top             =   3720
      Width           =   3255
   End
   Begin VB.CommandButton Cmdquit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   7
      Top             =   6960
      Width           =   1455
   End
   Begin VB.CommandButton Cmdbeauty 
      Caption         =   "Beauty and the Beast Summary"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9960
      TabIndex        =   6
      Top             =   2400
      Width           =   3255
   End
   Begin VB.CommandButton Cmdback 
      Caption         =   "Back to the Future Summary"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9960
      TabIndex        =   5
      Top             =   1320
      Width           =   3255
   End
   Begin VB.CommandButton Cmdalmost 
      Caption         =   "Almost Famous Summary"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9960
      TabIndex        =   4
      Top             =   240
      Width           =   3255
   End
   Begin VB.CommandButton Cmdform4 
      Caption         =   "Go to Beauty and the Beast Trivia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   3
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton Cmdform3 
      Caption         =   "Go to Back to the Future Trivia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   2
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton Cmdform2 
      Caption         =   "Go to Almost Famous Trivia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   1
      Top             =   3360
      Width           =   1455
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8295
      Left            =   3480
      ScaleHeight     =   8235
      ScaleWidth      =   6195
      TabIndex        =   0
      Top             =   120
      Width           =   6255
   End
End
Attribute VB_Name = "FrmForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Movie Trivia
'FrmMenu
'Amber Olson, Emily Borka, Shannon O'Neill
'11-1-08
'The purpose of this form is to allow users to get info about the movies and to navigate to each different page.
'The Purpose of this project was to create a movie trivia game that intigrated coded techniques discussed in lecture and lab
'and to make use of color and images.

Private Sub Cmdalmost_Click()
'This is to open where we have our text saved. and to print it into the picture box.
Dim Famous As String
Open App.Path & "\almostfamous.txt" For Input As #1
Do Until EOF(1)
    Input #1, Famous
    PicResults.Print Famous
Loop
Close #1

End Sub

Private Sub Cmdback_Click()
'This is to open where we have our text saved. and to print it into the picture box.
Dim backtothefuturetrivia As String
    Open App.Path & "\backtothefuturetrivia.txt" For Input As #1
    Do Until EOF(1)
    Input #1, backtothefuturetrivia
    PicResults.Print backtothefuturetrivia
Loop
Close #1
End Sub

Private Sub Cmdbeauty_Click()
'This is to open where we have our text saved. and to print it into the picture box.
Dim Beauty As String
Open App.Path & "\beautyandthebeast.txt" For Input As #1
Do Until EOF(1)
    Input #1, Beauty
    PicResults.Print Beauty
Loop
Close #1
End Sub

Private Sub cmdClear_Click()
'This is to clear the picture box.
    PicResults.Cls
    
End Sub

Private Sub Cmdform2_Click()
'This is to get to the Almost Famous page, and hide other pages.
FrmForm1.Hide
FrmForm2.Show
FrmForm3.Hide
FrmForm4.Hide
Frmsearch.Hide



End Sub

Private Sub Cmdform3_Click()
'This is to get to the Back to the Future page, and hide other pages.
FrmForm1.Hide
FrmForm2.Hide
FrmForm3.Show
FrmForm4.Hide
Frmsearch.Hide
End Sub

Private Sub Cmdform4_Click()
'This is to get to the Beauty and the Beast page, and hide other pages.
FrmForm1.Hide
FrmForm2.Hide
FrmForm3.Hide
FrmForm4.Show
Frmsearch.Hide
End Sub

Private Sub Cmdquit_Click()
'This is end the program.

End
End Sub

Private Sub Cmdsearch_Click()
'This is to get to the Search page, and hide other pages.
FrmForm1.Hide
FrmForm2.Hide
FrmForm3.Hide
FrmForm4.Hide
Frmsearch.Show


End Sub

Private Sub Form_Load()
'This is to get the users name to use later in the program.
    UserName = InputBox("Please Enter Your Name:", "What Is Your Name?")
End Sub
