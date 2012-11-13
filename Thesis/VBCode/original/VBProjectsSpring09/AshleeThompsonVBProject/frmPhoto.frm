VERSION 5.00
Begin VB.Form frmPhoto 
   BackColor       =   &H00000080&
   Caption         =   "Photography"
   ClientHeight    =   10935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15225
   LinkTopic       =   "Form1"
   ScaleHeight     =   10935
   ScaleWidth      =   15225
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Previous"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   960
      TabIndex        =   4
      Top             =   9600
      Width           =   1575
   End
   Begin VB.CommandButton cmdMain 
      Caption         =   "Back to Main Menu"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   960
      TabIndex        =   3
      Top             =   8520
      Width           =   1575
   End
   Begin VB.CommandButton cmdDigital 
      Caption         =   "Digital Photography"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   840
      TabIndex        =   2
      Top             =   5760
      Width           =   2175
   End
   Begin VB.CommandButton cmdBalckWhite 
      Caption         =   "Black and White Photography"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   840
      TabIndex        =   1
      Top             =   1080
      Width           =   2175
   End
   Begin VB.PictureBox picPhoto 
      Height          =   9495
      Left            =   3960
      Picture         =   "frmPhoto.frx":0000
      ScaleHeight     =   9435
      ScaleWidth      =   9435
      TabIndex        =   0
      Top             =   600
      Width           =   9495
   End
End
Attribute VB_Name = "frmPhoto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'The Artist's Multimedia Portfolio
'frmPhoto
'Ashley Thompson
'Friday March 20, 2009
'This form using file input and reads files containing the artist's photography into arrays
'and then shows a slideshow of the jpg images in a picture box using the Timer function
'It also conatins a button that takes the user back to the previous form and one to take the user to the main menu

Private Sub cmdBack_Click()
frmPhoto.Hide
frmArtMain.Show
End Sub

Private Sub cmdBalckWhite_Click()

Dim BWPhoto(1 To 6) As String
Dim ctr3 As Integer

Open App.Path & "\BWPhotos.txt" For Input As #6

ctr3 = 0

Do While Not EOF(6)
  ctr3 = ctr3 + 1
    Input #6, BWPhoto(ctr3)
Loop

Close #6

Dim whichOne As Integer, stopper As Integer, t As Double, oldOne As Integer, ctr4 As Double


whichOne = 1

stopper = 0

Do While (stopper < 6)

    picPhoto.Picture = LoadPicture(App.Path & "\" & BWPhoto(whichOne))
 
    
    
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



Private Sub cmdDigital_Click()
Dim DigiPhoto(1 To 10) As String
Dim ctr5 As Integer

Open App.Path & "\DigiPhotos.txt" For Input As #7

ctr5 = 0

Do While Not EOF(7)
    ctr5 = ctr5 + 1
    Input #7, DigiPhoto(ctr5)
Loop

Close #7

Dim whichOne As Integer, stopper As Integer, t As Double, oldOne As Integer, ctr6 As Double


whichOne = 1

stopper = 0

Do While (stopper < 10)

    picPhoto.Picture = LoadPicture(App.Path & "\" & DigiPhoto(whichOne))
 
    
    
    t = Timer
    Do While (Timer - t) < 1
        ctr6 = ctr6 + 1
        If ctr5 = 1000000 Then
            
            ctr6 = 0
        End If
    Loop
    
   
    stopper = stopper + 1
    
    oldOne = whichOne
    whichOne = (stopper Mod ctr6) + 1
    
Loop

End Sub

Private Sub cmdMain_Click()
frmMain.Show
frmPhoto.Hide
End Sub
