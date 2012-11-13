VERSION 5.00
Begin VB.Form characters 
   BackColor       =   &H80000001&
   Caption         =   "Form1"
   ClientHeight    =   8835
   ClientLeft      =   660
   ClientTop       =   450
   ClientWidth     =   10995
   LinkTopic       =   "Form1"
   ScaleHeight     =   8835
   ScaleWidth      =   10995
   Begin VB.CommandButton Command3 
      Caption         =   "Quit"
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   8160
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Use Character"
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   5160
      Width           =   1575
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFFFF&
      Height          =   8655
      Left            =   2280
      ScaleHeight     =   8595
      ScaleWidth      =   8595
      TabIndex        =   1
      Top             =   0
      Width           =   8655
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000012&
      Caption         =   "View Character"
      Height          =   615
      Left            =   240
      MaskColor       =   &H80000013&
      TabIndex        =   0
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000013&
      Caption         =   $"start.frx":0000
      Height          =   2415
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   1575
   End
End
Attribute VB_Name = "characters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Super Smash Bros.
'Opening Form
'Ryan Poster and Erik Skoe
'March 26th
'The object is to display the characters.



Private Sub Command1_Click()
Dim found As Boolean
found = False

picResults.Cls

Do While Not found  'To let the user choose a character to view.
Char = InputBox("Please Enter Number")
If Char > 12 Or Char <= 0 Then
MsgBox ("Please enter a value from the table above.")
Else: picResults.Picture = LoadPicture(App.Path & "\" & Pics(Char))
found = True
End If
Loop

End Sub

Private Sub Command2_Click()
If Char = 0 Then
MsgBox ("Please select a character.")
Else:   characters.Hide 'To hide the character form and show the damage form.
        damage.Show
End If
End Sub

Private Sub Command3_Click()
End 'To end the programm.
End Sub


Private Sub Form_Load()        'This will load a background picture for the characters form.
picResults.Picture = LoadPicture("selection.jpg")
End Sub
