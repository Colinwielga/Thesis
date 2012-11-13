VERSION 5.00
Begin VB.Form frmtrivia 
   Caption         =   "Form1"
   ClientHeight    =   7935
   ClientLeft      =   5565
   ClientTop       =   3855
   ClientWidth     =   10875
   LinkTopic       =   "Form1"
   Picture         =   "frmtrivia.frx":0000
   ScaleHeight     =   7935
   ScaleWidth      =   10875
   Begin VB.CommandButton cmdread 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Read the File"
      Height          =   615
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2640
      Width           =   1695
   End
   Begin VB.PictureBox picresults 
      Height          =   1455
      Left            =   2880
      ScaleHeight     =   1395
      ScaleWidth      =   5355
      TabIndex        =   4
      Top             =   4800
      Width           =   5415
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Next"
      Enabled         =   0   'False
      Height          =   615
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3480
      Width           =   1695
   End
   Begin VB.CommandButton cmdPrevious 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Previous"
      Enabled         =   0   'False
      Height          =   615
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3480
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdhome 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Back to Neverland!"
      Height          =   735
      Left            =   8280
      MaskColor       =   &H0080FF80&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6960
      Width           =   2175
   End
   Begin VB.Label lbltitle 
      BackStyle       =   0  'Transparent
      Caption         =   "DID YOU KNOW..."
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   2295
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   8535
   End
End
Attribute VB_Name = "frmtrivia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The Wonderful World of Disney
'form Home
'Kate Hendrickx
'February 2010
'Objective: to teach the user something new.

Option Explicit
Dim CTR As Single, trivia(1 To 4) As String, ctr2 As Single

Private Sub cmdhome_Click()
frmtrivia.Hide
frmhome.Show
End Sub

Private Sub cmdNext_Click()

'printing trivia one at a time going forward.
If ctr2 <= 3 Then
ctr2 = ctr2 + 1
Else: ctr2 = 1
End If
picresults.Cls
picresults.Print trivia(ctr2)

'after next is clicked for the first time, the button 'previous' is enabled.
cmdPrevious.Visible = True

End Sub

Private Sub cmdPrevious_Click()

'printing trivia one at a time in reverse.
If ctr2 >= 2 Then
ctr2 = ctr2 - 1
Else: ctr2 = 4
End If
picresults.Cls
picresults.Print trivia(ctr2)

End Sub

Private Sub cmdread_Click()

'opening the file
Open App.Path & "\trivia.txt" For Input As #1


'reading the array
CTR = 0
Do While Not EOF(1)
CTR = CTR + 1
Input #1, trivia(CTR)
Loop
Close #1

ctr2 = 0

'enabling the other buttons
cmdNext.Enabled = True
cmdPrevious.Enabled = True
End Sub
