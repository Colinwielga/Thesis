VERSION 5.00
Begin VB.Form frmMadrid 
   BackColor       =   &H00FF80FF&
   Caption         =   "Form1"
   ClientHeight    =   7215
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9180
   LinkTopic       =   "Form1"
   ScaleHeight     =   7215
   ScaleWidth      =   9180
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FFC0FF&
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
      BackColor       =   &H00FF80FF&
      Height          =   4935
      Left            =   120
      ScaleHeight     =   4875
      ScaleWidth      =   5835
      TabIndex        =   4
      Top             =   120
      Width           =   5895
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FFC0FF&
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
      BackColor       =   &H00FFC0FF&
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
      BackColor       =   &H00FFC0FF&
      Caption         =   "Find Attractions In Madrid"
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
      BackColor       =   &H00FFC0FF&
      Caption         =   "View Slide Show of Madrid"
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
      BackColor       =   &H00FF80FF&
      Caption         =   "Visiting Madrid!"
      BeginProperty Font 
         Name            =   "MS PMincho"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   1575
      Left            =   1920
      TabIndex        =   6
      Top             =   5400
      Width           =   2895
   End
End
Attribute VB_Name = "frmMadrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FinalProject:Travel Europe
'frmMadrid
'Jessica Florek
'Written: 3/10/09
'Objective: This form displays a slide show of pictures of Madrid as well as has
'button the user can select to book a hotel and find entertinmetn in Madrid.


Option Explicit

'this slide acts as a navigation device through the different activites that the user can accomplish relating the the city of Madrid
Private Sub cmdAttractions_Click()
frmMadrid.Hide
frmMadridAttractions.Show

End Sub

Private Sub cmdBack_Click()
frmMadrid.Hide
frmMapCities.Show

End Sub

Private Sub cmdHotel_Click()
frmMadrid.Hide
frmMadridHotel.Show
cmdHotel.Enabled = False

End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdSlideshow_Click()
Dim whichone As Integer, stopper As Integer, t As Double, oldone As Integer, ctr2 As Double, ctr As Integer
Dim MadridArray(1 To 10) As String

Open App.Path & "\MadridPics\MadridPicsArray.txt" For Input As #1
'loads picture names into an array that will be used to display the pictures in the slide show
Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, MadridArray(ctr)
Loop
Close #1

whichone = 1
stopper = 0

'slide show loops
Do While (stopper < ctr)
    picResults.Picture = LoadPicture(App.Path & "\MadridPics\" & MadridArray(whichone))
    
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
