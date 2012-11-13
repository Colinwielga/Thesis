VERSION 5.00
Begin VB.Form frmLondon 
   BackColor       =   &H0080C0FF&
   Caption         =   "Form1"
   ClientHeight    =   7965
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8865
   LinkTopic       =   "Form1"
   ScaleHeight     =   7965
   ScaleWidth      =   8865
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSlideshow 
      BackColor       =   &H00C0E0FF&
      Caption         =   "View Slide Show of London"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   480
      Width           =   1935
   End
   Begin VB.CommandButton cmdAttractions 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Find Attractions In London"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2760
      Width           =   1935
   End
   Begin VB.CommandButton cmdHotel 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Book a Hotel"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1560
      Width           =   1935
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5160
      Width           =   1935
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H0080C0FF&
      Height          =   4935
      Left            =   360
      ScaleHeight     =   4875
      ScaleWidth      =   5835
      TabIndex        =   1
      Top             =   1200
      Width           =   5895
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Go Back "
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H0080C0FF&
      Caption         =   "Visiting London!"
      BeginProperty Font 
         Name            =   "Centaur"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   720
      TabIndex        =   5
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "frmLondon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FinalProject:Travel Europe
'frmLondon
'Jessica Florek
'Written: 3/6/09
'Objective: This form displays a slide show of pictures of London as well as has
'button the user can select to book a hotel and find entertinmetn in London.

Option Explicit

'this form acts as a navigator to different forms to complete various options
Private Sub cmdAttractions_Click()
frmLondon.Hide
frmLondonAttractions.Show
End Sub

Private Sub cmdBack_Click()
frmLondon.Hide
frmMapCities.Show
End Sub

Private Sub cmdHotel_Click()
frmLondon.Hide
frmLondonHotel.Show
cmdHotel.Enabled = False
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdSlideshow_Click()
Dim whichone As Integer, stopper As Integer, t As Double, oldone As Integer, ctr2 As Double, ctr As Integer
Dim LondonArray(1 To 10) As String

Open App.Path & "\LondonPics\LondonPicsArray.txt" For Input As #1
'creates an array that will later tell the slide show what pictures to display
Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, LondonArray(ctr)
Loop
Close #1

whichone = 1
stopper = 0

Do While (stopper < ctr)
    picResults.Picture = LoadPicture(App.Path & "\LondonPics\" & LondonArray(whichone))
    
    'Timer function dispays each picture for about 2 seconds ****this coding was taken and modified from VB examples folder****
     t = Timer
    Do While (Timer - t) < 2
        ctr2 = ctr2 + 1
        If ctr2 = 1000000 Then
            ctr2 = 0
        End If
    Loop
    
    stopper = stopper + 1
    oldone = whichone
    whichone = (stopper Mod ctr) + 1
    Close #1
Loop


End Sub


