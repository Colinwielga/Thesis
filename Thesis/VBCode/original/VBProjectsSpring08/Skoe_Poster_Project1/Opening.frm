VERSION 5.00
Begin VB.Form Opening 
   BackColor       =   &H00008000&
   Caption         =   "Form1"
   ClientHeight    =   8805
   ClientLeft      =   660
   ClientTop       =   450
   ClientWidth     =   10905
   LinkTopic       =   "Form1"
   ScaleHeight     =   8805
   ScaleWidth      =   10905
   Begin VB.CommandButton Command2 
      Caption         =   "Quit"
      Height          =   1095
      Left            =   7080
      TabIndex        =   2
      Top             =   5880
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "View Characters!"
      Height          =   1095
      Left            =   840
      TabIndex        =   1
      Top             =   5880
      Width           =   2655
   End
   Begin VB.PictureBox picResults3 
      BackColor       =   &H00FFFFFF&
      Height          =   5775
      Left            =   1080
      ScaleHeight     =   5715
      ScaleWidth      =   8475
      TabIndex        =   0
      Top             =   0
      Width           =   8535
   End
End
Attribute VB_Name = "Opening"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Super Smash Bros.
'Opening Form
'Ryan Poster and Erik Skoe
'March 27th
'The object is to start the program and display a begining picture

Private Sub Command1_Click()
characters.Show 'To go to character form
Opening.Hide

End Sub

Private Sub Command2_Click()
End 'To end the programm
End Sub

Private Sub Form_Load()
picResults3.Picture = LoadPicture("Start.jpg")  'To load a background picture
End Sub

