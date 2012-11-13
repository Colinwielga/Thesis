VERSION 5.00
Begin VB.Form frmVehicles 
   Caption         =   "Deer and Vehicles"
   ClientHeight    =   10800
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10800
   LinkTopic       =   "Form1"
   Picture         =   "frmVehicles.frx":0000
   ScaleHeight     =   10800
   ScaleWidth      =   10800
   Begin VB.PictureBox picVPics 
      BackColor       =   &H00C0FFFF&
      Height          =   5535
      Left            =   2280
      ScaleHeight     =   5475
      ScaleWidth      =   6195
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   6255
   End
   Begin VB.Timer tmrSlideShow 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   10320
      Top             =   4320
   End
   Begin VB.CommandButton cmdPics 
      BackColor       =   &H0000FFFF&
      Caption         =   "What can a deer do to a car? Look at these collision photos."
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2160
      Width           =   2055
   End
   Begin VB.CommandButton cmdHome 
      BackColor       =   &H0000FFFF&
      Caption         =   "Return to Home Page"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8760
      Width           =   2055
   End
   Begin VB.CommandButton cmdFacts 
      BackColor       =   &H0000FFFF&
      Caption         =   "A few facts about deer-vehicle collisions."
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   2055
   End
   Begin VB.PictureBox picVehicles 
      BackColor       =   &H0080C0FF&
      DrawWidth       =   5
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   2280
      ScaleHeight     =   4995
      ScaleWidth      =   6195
      TabIndex        =   3
      Top             =   5640
      Visible         =   0   'False
      Width           =   6255
   End
   Begin VB.CommandButton cmdHit 
      BackColor       =   &H0000FFFF&
      Caption         =   "What to do if you can't avoid hitting a deer."
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2280
      Width           =   2055
   End
   Begin VB.CommandButton cmdTips 
      BackColor       =   &H0000FFFF&
      Caption         =   "Tips for avoiding collisions with deer."
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H0000FFFF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8760
      Width           =   2055
   End
End
Attribute VB_Name = "frmVehicles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: MN Deer'
'Form Name: Startup'
'Authors: Jordon Przybilla'
'Date Written: October 4, 2009
'this form will deal with everything to do with deer-vehicle collisions

Option Explicit
Dim q As Integer

Private Sub cmdFacts_Click()
'this button will display a few facts about car-deer collisions in a picture box
'the array was read as this form was being opened

picVPics.Visible = False
picVehicles.Visible = True
picVehicles.Cls

For x = 1 To Ctr
    picVehicles.Print Facts(x)
    picVehicles.Print
Next x



End Sub

Private Sub cmdHit_Click()

'this button displays what a driver should do if they can not avoid a collision with a deer

picVPics.Visible = False
picVehicles.Visible = True
picVehicles.Cls

For x = 1 To Ctr
    picVehicles.Print CantAvoid(x)
    picVehicles.Print
Next x

End Sub

Private Sub cmdHome_Click() 'take user to home page

frmVehicles.Hide
frmHome.Show

End Sub


Private Sub cmdPics_Click() 'this button will start a slide show that shows pictures of vehicles that have been in collisions with deer
q = 1
picVPics.Visible = True
tmrSlideShow.Enabled = True
picVehicles.Visible = False

End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdTips_Click()
'this button will give the user tips on how to avoid collisions with deer when driving
picVPics.Visible = False
picVehicles.Visible = True
picVehicles.Cls

For x = 1 To Ctr
    picVehicles.Print AvoidTips(x)
    picVehicles.Print
Next x



End Sub

Private Sub Form_Load()
' this will load the array that is used for the slide show in this form



Open App.Path & "\Data\collisionslides.txt" For Input As #1
        Ctr = 0
        Do While Not EOF(1)
            Ctr = Ctr + 1
            Input #1, collisionslides(Ctr)
        Loop
    Close #1

End Sub

Private Sub tmrSlideShow_Timer() 'this is the timer for the slide show and loads the pictures for that slide show as needed


If q < 7 Then
    picVPics.Picture = LoadPicture(App.Path & "\Project Pics\" & collisionslides(q))
    q = q + 1
Else
    tmrSlideShow.Enabled = False
End If

End Sub
