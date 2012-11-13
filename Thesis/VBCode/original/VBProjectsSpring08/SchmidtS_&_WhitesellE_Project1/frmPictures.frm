VERSION 5.00
Begin VB.Form frmPictures 
   BackColor       =   &H0080FFFF&
   Caption         =   "Study Abroad Pictures"
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10320
   LinkTopic       =   "Form1"
   ScaleHeight     =   6810
   ScaleWidth      =   10320
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGetStarted 
      BackColor       =   &H0080FF80&
      Caption         =   "Check out programs"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton cmdViewShow 
      BackColor       =   &H008080FF&
      Caption         =   "Click to View A Picture"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   3015
   End
   Begin VB.PictureBox picPictures 
      BackColor       =   &H00FFFFFF&
      Height          =   5175
      Left            =   1560
      ScaleHeight     =   5115
      ScaleWidth      =   8355
      TabIndex        =   0
      Top             =   1320
      Width           =   8415
   End
End
Attribute VB_Name = "frmPictures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'written 3/25/08 by Sammi and Erika

Private Sub cmdViewShow_Click()
Dim Which As Integer, Stopper As Integer, T As Double, OldOne As Integer

Which = 1

Stopper = 0

Do While (Stopper < 30)
    picResults.Picture = LoadPicture(App.Path & "\" & Name(Which))
    
    T = Timer
    Do While (Timer - T) < 3
    Loop
    
    Stopper = Stopper + 1
    
    OldOne = Which
    Which = (Stopper Mod CTR) + 1
Loop




End Sub



Private Sub cmdGetStarted_Click()
frmPictures.Hide
frmPrograms.Show
End Sub

