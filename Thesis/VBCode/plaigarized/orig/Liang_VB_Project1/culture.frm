VERSION 5.00
Begin VB.Form frmculture 
   BackColor       =   &H0000FFFF&
   Caption         =   "Form1"
   ClientHeight    =   11340
   ClientLeft      =   3225
   ClientTop       =   3135
   ClientWidth     =   21765
   FillColor       =   &H000080FF&
   ForeColor       =   &H000080FF&
   LinkTopic       =   "Form1"
   Picture         =   "culture.frx":0000
   ScaleHeight     =   11340
   ScaleWidth      =   21765
   Begin VB.CommandButton cmdhome 
      Caption         =   "Homepage"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   17520
      TabIndex        =   3
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Timer tmrSlideShow 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   16200
      Top             =   1560
   End
   Begin VB.PictureBox picResults1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8775
      Left            =   12120
      ScaleHeight     =   8715
      ScaleWidth      =   9315
      TabIndex        =   2
      Top             =   2160
      Width           =   9375
   End
   Begin VB.CommandButton cmdshow 
      Caption         =   "Click Here!"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12360
      TabIndex        =   1
      Top             =   1440
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   $"culture.frx":1E756
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   12000
      TabIndex        =   0
      Top             =   240
      Width           =   9015
   End
End
Attribute VB_Name = "frmculture"
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
Dim Culture(1 To 10) As String
Dim CTR As Integer
Dim PicIndex As Integer


Private Sub cmdhome_Click()
frmculture.Hide
frmMain.Show
End Sub

Private Sub cmdshow_Click()
'Sets the variable PicIndex to one
PicIndex = 1
'Enables the timer button which starts the slide show
tmrSlideShow.Enabled = True
End Sub

Private Sub tmrSlideShow_Timer()
'starts the slide show when enabled by the slideshow button
'the if statement makes sure that the PicIndex doesn't go past the amount in the array
    If PicIndex <= 7 Then
    'clears the picture box for the titles
        picResults1.Cls
    'displays the pictures
        picResults1.Picture = LoadPicture(App.Path & "\Image1\" & Culture(PicIndex))
    'displays the title of the picture displayed
        picResults1.Print Culture(PicIndex)
        PicIndex = PicIndex + 1
    Else
        tmrSlideShow.Enabled = False
    End If
End Sub

Private Sub Form_Load()
'Initializes CTR
CTR = 0
'Loads the file "Tourism.txt" into an array Tourism
Open App.Path & "\Culture.txt" For Input As #1
Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, Culture(CTR)
Loop
Close #1

End Sub
