VERSION 5.00
Begin VB.Form frmVenice 
   BackColor       =   &H00C0C000&
   Caption         =   "Form1"
   ClientHeight    =   6150
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8580
   LinkTopic       =   "Form1"
   ScaleHeight     =   6150
   ScaleWidth      =   8580
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSlideshow 
      BackColor       =   &H0080C0FF&
      Caption         =   "View Slide Show of Venice"
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
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton cmdAttractions 
      BackColor       =   &H0080C0FF&
      Caption         =   "Find Attractions In Venice"
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
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2520
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
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1320
      Width           =   1935
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
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4920
      Width           =   1935
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00C0C000&
      Height          =   4335
      Left            =   240
      ScaleHeight     =   4275
      ScaleWidth      =   4875
      TabIndex        =   1
      Top             =   840
      Width           =   4935
   End
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
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00C0C000&
      Caption         =   "Visiting Venice!"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   615
      Left            =   1320
      TabIndex        =   6
      Top             =   0
      Width           =   3255
   End
End
Attribute VB_Name = "frmVenice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FinalProject:Travel Europe
'frmVenice
'Jessica Florek
'Written: 3/12/09
'Objective: This form displays a slide show of pictures of Venice as well as has
'button the user can select to book a hotel and find entertinmetn in Venice.

Option Explicit

'navigates user through the options corresponding to visiting venice ie. booking hotel and entertainment options
Private Sub cmdAttractions_Click()
frmVenice.Hide
frmVeniceAttractions.Show

End Sub

Private Sub cmdBack_Click()
frmVenice.Hide
frmMapCities.Show
End Sub

Private Sub cmdHotel_Click()
frmVenice.Hide
frmVeniceHotel.Show
cmdHotel.Enabled = False

End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdSlideshow_Click()
Dim whichone As Integer, stopper As Integer, t As Double, oldone As Integer, ctr2 As Double, ctr As Integer
Dim VeniceArray(1 To 10) As String

Open App.Path & "\VenicePics\VenicePicsArray.txt" For Input As #1
'loads picture names into an array to use for the slide show
Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, VeniceArray(ctr)
Loop
Close #1

whichone = 1
stopper = 0
'slide show
Do While (stopper < ctr)
    picResults.Picture = LoadPicture(App.Path & "\VenicePics\" & VeniceArray(whichone))
    
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
