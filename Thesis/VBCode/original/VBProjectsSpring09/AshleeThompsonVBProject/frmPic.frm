VERSION 5.00
Begin VB.Form frmPic 
   BackColor       =   &H00008080&
   Caption         =   "Artist's Information"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   13875
   LinkTopic       =   "Form2"
   ScaleHeight     =   11010
   ScaleWidth      =   13875
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   3375
      Left            =   120
      Picture         =   "frmPic.frx":0000
      ScaleHeight     =   3315
      ScaleWidth      =   2595
      TabIndex        =   5
      Top             =   120
      Width           =   2655
   End
   Begin VB.CommandButton cmdBackground 
      Caption         =   "Background"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   480
      TabIndex        =   4
      Top             =   6480
      Width           =   1575
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Main Menu"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   12000
      TabIndex        =   3
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton cmdResume 
      Caption         =   "Artist's Resume"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   480
      TabIndex        =   2
      Top             =   5160
      Width           =   1575
   End
   Begin VB.PictureBox picResults 
      Height          =   9255
      Left            =   3000
      ScaleHeight     =   9195
      ScaleWidth      =   10515
      TabIndex        =   1
      Top             =   2280
      Width           =   10575
   End
   Begin VB.CommandButton cmdStatement 
      Caption         =   "Artist's Statement"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   480
      TabIndex        =   0
      Top             =   3840
      Width           =   1575
   End
End
Attribute VB_Name = "frmPic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'The Artist's Multimedia Portfolio
'frmPic
'Ashley Thompson
'Friday March 20, 2009
'This form uses file input to read three files, one containing the artist's background information,
'the next containing an artist's statement and the last containting the artist's resume
'It also uses a commmand button to take the user back to the main menu form


Private Sub cmdBackground_Click()

picResults.Cls

Dim AName(1 To 2) As String
Dim aTitle(1 To 2) As String
Dim Background(1 To 2) As String


Open App.Path & "\Background.txt" For Input As #4

ctr = 0

Do While Not EOF(4)
  ctr = ctr + 1
    Input #4, AName(ctr), aTitle(ctr), Background(ctr)
    
    picResults.Print AName(ctr)
    picResults.Print aTitle(ctr)
    picResults.Print "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
    picResults.Print Background(ctr)
   
Loop
Close #4
End Sub



Private Sub cmdResume_Click()
picResults.Cls

Dim Resum(1 To 2) As String

Open App.Path & "\Resume.txt" For Input As #3

ctr = 0

Do While Not EOF(3)
  ctr = ctr + 1
    Input #3, Resum(ctr)
    
    picResults.Print Resum(ctr)
    
Loop
Close #3
End Sub

Private Sub cmdReturn_Click()
frmPic.Hide
frmMain.Show
End Sub

Private Sub cmdStatement_Click()
picResults.Cls

Dim Artist(1 To 2) As String
Dim Title(1 To 2) As String
Dim Statement(1 To 2) As String

Open App.Path & "\ArtistStatement.txt" For Input As #2

ctr = 0

Do While Not EOF(2)
  ctr = ctr + 1
    Input #2, Artist(ctr), Title(ctr), Statement(ctr)
    
    picResults.Print Artist(ctr)
    picResults.Print Title(ctr)
    picResults.Print "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
    picResults.Print Statement(ctr)
Loop
Close #2
    
 

End Sub


