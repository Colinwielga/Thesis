VERSION 5.00
Begin VB.Form frmSketch 
   BackColor       =   &H00400040&
   Caption         =   "Sketch Book"
   ClientHeight    =   9855
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13530
   LinkTopic       =   "Form1"
   ScaleHeight     =   9855
   ScaleWidth      =   13530
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMain 
      Caption         =   "Back to Main Menu"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   960
      TabIndex        =   3
      Top             =   6360
      Width           =   1815
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Previous"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   960
      TabIndex        =   2
      Top             =   4920
      Width           =   1815
   End
   Begin VB.CommandButton cmdSketchBk 
      Caption         =   "View Sketch Book"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   720
      TabIndex        =   1
      Top             =   600
      Width           =   2055
   End
   Begin VB.PictureBox picSketch 
      Height          =   9135
      Left            =   4200
      Picture         =   "frmSketch.frx":0000
      ScaleHeight     =   9075
      ScaleWidth      =   9075
      TabIndex        =   0
      Top             =   360
      Width           =   9135
   End
End
Attribute VB_Name = "frmSketch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'The Artist's Multimedia Portfolio
'frmSketch
'Ashley Thompson
'Friday March 20, 2009
'This form using file input and reads files containing the artist's sketches into arrays
'and then shows a slideshow of the jpg images in a picture box using the Timer function
'It also conatins a button that takes the user back to the previous form and one to take the user to the main menu

Private Sub cmdBack_Click()
frmSketch.Hide
frmArtMain.Show
End Sub



Private Sub cmdMain_Click()
frmSketch.Hide
frmMain.Show
End Sub

Private Sub cmdSketchBk_Click()

Dim Sketch(1 To 20) As String
Dim ctra As Integer

Open App.Path & "\Sketchbook.txt" For Input As #8

ctra = 0

Do While Not EOF(8)
  ctra = ctra + 1
    Input #8, Sketch(ctra)
Loop

Close #8

Dim whichOne As Integer, stopper As Integer, t As Double, oldOne As Integer, ctr4 As Double


whichOne = 1

stopper = 0

Do While (stopper < 20)

    picSketch.Picture = LoadPicture(App.Path & "\" & Sketch(whichOne))
 
    
    
    t = Timer
    Do While (Timer - t) < 1
        ctr4 = ctr4 + 1
        If ctr4 = 1000000 Then
            
            ctr4 = 0
        End If
    Loop
    
   
    stopper = stopper + 1
    
    oldOne = whichOne
    whichOne = (stopper Mod ctr4) + 1
    
Loop

End Sub

Private Sub Command1_Click()
frmSketch.Hide
frmMain.Show
End Sub
