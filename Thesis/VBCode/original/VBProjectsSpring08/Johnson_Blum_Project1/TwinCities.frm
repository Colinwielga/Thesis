VERSION 5.00
Begin VB.Form TwinCities 
   BackColor       =   &H000000C0&
   Caption         =   "Form1"
   ClientHeight    =   6660
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8955
   BeginProperty Font 
      Name            =   "Kristen ITC"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6660
   ScaleWidth      =   8955
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdback 
      Caption         =   "Homepage"
      Height          =   855
      Left            =   1080
      TabIndex        =   8
      Top             =   5640
      Width           =   1575
   End
   Begin VB.PictureBox picResultsTitleName 
      Height          =   375
      Left            =   3600
      ScaleHeight     =   315
      ScaleWidth      =   4215
      TabIndex        =   7
      Top             =   6120
      Width           =   4275
   End
   Begin VB.Timer tmrSlideShow 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   600
      Top             =   5760
   End
   Begin VB.CommandButton cmdSlideShow 
      BackColor       =   &H00FF0000&
      Caption         =   "Watch a Slide Show"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4560
      Width           =   2655
   End
   Begin VB.CommandButton cmdView 
      BackColor       =   &H00FF0000&
      Caption         =   "<=Enter Integer and View Picture"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox txtPicture 
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
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   735
   End
   Begin VB.PictureBox picResults 
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   2880
      ScaleHeight     =   4635
      ScaleWidth      =   5595
      TabIndex        =   0
      Top             =   1440
      Width           =   5655
   End
   Begin VB.Label lblOr 
      BackColor       =   &H000000C0&
      Caption         =   "Or"
      BeginProperty Font 
         Name            =   "Matura MT Script Capitals"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   600
      TabIndex        =   5
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Twin Cities Tourist Attractions"
      BeginProperty Font 
         Name            =   "Mathematica6"
         Size            =   36
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   2400
      TabIndex        =   2
      Top             =   0
      Width           =   6255
   End
   Begin VB.Label lblPictures 
      BackColor       =   &H00FF0000&
      Caption         =   $"TwinCities.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2055
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "TwinCities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Minnesoooota
'Form Name: TwinCities
'Author: Danielle Johnson and Tony Blum
'Date Written: March 26th 2008
'The purpose of this form is to allow the user to view pictures of tourist attractions around the Twin Cities
Option Explicit
'Declares variables globally throughout the form
Dim Tourism(1 To 10) As String
Dim CTR As Integer
Dim PicIndex As Integer

Private Sub cmdback_Click()
TwinCities.Hide 'hide this form, go back to homepage
Minnesota.Show
End Sub

Private Sub Form_Load()
'Initializes CTR
CTR = 0
'Loads the file "Tourism.txt" into an array Tourism
Open App.Path & "\Tourism.txt" For Input As #1
Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, Tourism(CTR)
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
'prints the results
picResults.Picture = LoadPicture(App.Path & "\" & Tourism(Picture))
End Sub


Private Sub tmrSlideShow_Timer()
'starts the slide show when enabled by the slideshow button
'the if statement makes sure that the PicIndex doesn't go past the amount in the array
    If PicIndex < 11 Then
    'clears the picture box for the titles
        picResultsTitleName.Cls
    'displays the pictures
        picResults.Picture = LoadPicture(App.Path & "\" & Tourism(PicIndex))
    'displays the title of the picture displayed
        picResultsTitleName.Print Tourism(PicIndex)
        PicIndex = PicIndex + 1
    Else
        tmrSlideShow.Enabled = False
    End If
        
End Sub
