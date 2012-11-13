VERSION 5.00
Begin VB.Form stadiums 
   BackColor       =   &H000000FF&
   Caption         =   "Form2"
   ClientHeight    =   6945
   ClientLeft      =   2025
   ClientTop       =   1620
   ClientWidth     =   11295
   LinkTopic       =   "Form2"
   ScaleHeight     =   6945
   ScaleWidth      =   11295
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF0000&
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5400
      Width           =   2775
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   1
      Left            =   10320
      Picture         =   "stadiums.frx":0000
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   8
      Top             =   6120
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   0
      Left            =   120
      Picture         =   "stadiums.frx":0752
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   7
      Top             =   6120
      Width           =   855
   End
   Begin VB.PictureBox picResultsTitleName 
      BackColor       =   &H00FF8080&
      Height          =   375
      Left            =   1440
      ScaleHeight     =   315
      ScaleWidth      =   4215
      TabIndex        =   6
      Top             =   6120
      Width           =   4275
   End
   Begin VB.Timer tmrSlideShow 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   10920
      Top             =   4560
   End
   Begin VB.CommandButton cmdSlideShow 
      BackColor       =   &H00FF0000&
      Caption         =   "Slide Show"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4320
      Width           =   2655
   End
   Begin VB.TextBox txtPicture 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Gungsuh"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      TabIndex        =   3
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton cmdView 
      BackColor       =   &H00FF0000&
      Caption         =   "<=Enter Number to View Picture"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2400
      Width           =   1695
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FF8080&
      Height          =   6015
      Left            =   0
      ScaleHeight     =   5955
      ScaleWidth      =   7995
      TabIndex        =   0
      Top             =   0
      Width           =   8055
   End
   Begin VB.Label lblOr 
      BackColor       =   &H000000FF&
      Caption         =   "Or"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9120
      TabIndex        =   4
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label lblPictures 
      BackColor       =   &H00FF0000&
      Caption         =   $"stadiums.frx":0EA4
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2055
      Left            =   8280
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "stadiums"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Declares variables globally throughout the form
Dim metropic(1 To 10) As String, pic(1 To 10) As String
Dim CTR As Integer
Dim PicIndex As Integer
'hides stadium form and shows history form
Private Sub cmdback_Click()
stadiums.Hide 'hide this form, go back to homepage
history.Show
End Sub
'hides stadium form and shows history form
Private Sub Command1_Click()
stadiums.Hide
history.Show

End Sub

Private Sub Form_Load()
'Initializes CTR
CTR = 0
'Loads the file "Tourism.txt" into an array Tourism
Open App.Path & "\metropics.txt" For Input As #1
Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, metropic(CTR)
   
Loop
Close #1

End Sub
Private Sub cmdSlideShow_Click()
'Sets the variable PicIndex to one
PicIndex = 1
'Enables the timer button which starts the slide show
tmrSlideShow.Enabled = True
End Sub

Private Sub cmdView_Click()
'Declares variables
Dim Picture As Integer
'Gives picture a value inputted by the user in a text box
Picture = txtPicture.Text
'Ensures that if the user enters a number that is not in the array that the program will continue to work
If Picture < 1 Or Picture > 10 Then
    MsgBox "The program only calculates integers between 1 and 10.  Look at the table above for help and enjoy this randomized photo!", , "Incorrect Number"
    Picture = (CInt(Int((10 * Rnd()) + 1)))
End If
'prints the result
picResults.Print metropic(Picture)
picResults.Picture = LoadPicture(App.Path & "\" & metropic(Picture))
End Sub


Private Sub tmrSlideShow_Timer()

    If PicIndex <= 6 Then
    
        picResultsTitleName.Cls
    
        picResults.Picture = LoadPicture(App.Path & "\" & metropic(PicIndex))
    
        picResultsTitleName.Print metropic(PicIndex)
        PicIndex = PicIndex + 1
    Else
        tmrSlideShow.Enabled = False
    End If
        
End Sub

