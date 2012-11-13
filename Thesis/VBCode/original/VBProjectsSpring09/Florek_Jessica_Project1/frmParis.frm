VERSION 5.00
Begin VB.Form frmParis 
   BackColor       =   &H000000C0&
   Caption         =   "Form1"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8580
   LinkTopic       =   "Form1"
   ScaleHeight     =   6435
   ScaleWidth      =   8580
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H0080C0FF&
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
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3840
      Width           =   1935
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H000000C0&
      Height          =   4335
      Left            =   240
      ScaleHeight     =   4275
      ScaleWidth      =   4875
      TabIndex        =   4
      Top             =   960
      Width           =   4935
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H0080C0FF&
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
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5040
      Width           =   1935
   End
   Begin VB.CommandButton cmdHotel 
      BackColor       =   &H0080C0FF&
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
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1440
      Width           =   1935
   End
   Begin VB.CommandButton cmdAttractions 
      BackColor       =   &H0080C0FF&
      Caption         =   "Find Attractions In Paris"
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
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2640
      Width           =   1935
   End
   Begin VB.CommandButton cmdSlideshow 
      BackColor       =   &H0080C0FF&
      Caption         =   "View Slide Show of Paris"
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
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H000000C0&
      Caption         =   "Visiting Paris!"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   615
      Left            =   1560
      TabIndex        =   6
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "frmParis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FinalProject:Travel Europe
'frmParis
'Jessica Florek
'Written: 3/6/09
'Objective: This form displays a slide show of pictures of Paris as well as has
'button the user can select to book a hotel and find entertinmetn in Paris.
Option Explicit

'navigates around the options that are related to visiting the city
Private Sub cmdAttractions_Click()
frmParis.Hide
frmParisAttractions.Show
End Sub

Private Sub cmdBack_Click()
frmParis.Hide
frmMapCities.Show
End Sub

Private Sub cmdHotel_Click()
frmParis.Hide
frmParisHotel.Show
cmdHotel.Enabled = False
'deactivates the hotel command so the user doesn't book the same hotel many times
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdSlideshow_Click()
Dim whichone As Integer, stopper As Integer, t As Double, oldone As Integer, ctr2 As Double, ctr As Integer
Dim ParisArray(1 To 10) As String
'loads list of picture names into an array to be used later in the slide show
Open App.Path & "\ParisPics\ParisPicsArray.txt" For Input As #1

Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, ParisArray(ctr)
Loop
Close #1

whichone = 1
stopper = 0
'slide show
Do While (stopper < ctr)
    picResults.Picture = LoadPicture(App.Path & "\ParisPics\" & ParisArray(whichone))
    
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

