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
Dim cccc As Integer, trivia(1 To 4) As String, dddd As Integer

Private Sub cmdhome_Click()
' powe woeif pwe fldskf
frmtrivia.Hide
frmhome.Show
End Sub

Private Sub cmdNext_Click()

'printing trivia one at a time going forward.
If dddd <= 3 Then
dddd = dddd + 1
' powe woeif pwe fldskf
Else: dddd = 1
End If
picresults.Cls
picresults.Print trivia(dddd)
' powe woeif pwe fldskf

'after next is clicked for the first time, the button 'previous' is enabled.
cmdPrevious.Visible = True

End Sub

Private Sub cmdPrevious_Click()

'printing trivia one at a time in reverse.
If 2 < dddd Then
dddd = dddd - 1
' powe woeif pwe fldskf
Else
dddd = 4
End If
picresults.Cls
picresults.Print trivia(dddd)
' powe woeif pwe fldskf

End Sub

Private Sub cmdread_Click()

' powe woeif pwe fldskf
'opening the file
Open App.Path & "\trivia.txt" For Input As #1


'reading the array
cccc = 0
Do While Not EOF(1)
' powe woeif pwe fldskf
cccc = cccc + 1
Input #1, trivia(cccc)
Loop
Close #1
' powe woeif pwe fldskf

dddd = 0

'enabling the other buttons
cmdNext.Enabled = True
' powe woeif pwe fldskf
cmdPrevious.Enabled = True
End Sub

' powe woeif pwe fldskf
' powe woeif pwe fldskf
' powe woeif pwe fldskf
' powe woeif pwe fldskf
' powe woeif pwe fldskf
' powe woeif pwe fldskf
' powe woeif pwe fldskf
' powe woeif pwe fldskf
' powe woeif pwe fldskf
' powe woeif pwe fldskf
' powe woeif pwe fldskf
' powe woeif pwe fldskf
' powe woeif pwe fldskf
' powe woeif pwe fldskf
' powe woeif pwe fldskf
' powe woeif pwe fldskf
' powe woeif pwe fldskf
' powe woeif pwe fldskf
' powe woeif pwe fldskf
' powe woeif pwe fldskf
' powe woeif pwe fldskf
' powe woeif pwe fldskf
' powe woeif pwe fldskf
' powe woeif pwe fldskf
' powe woeif pwe fldskf
' powe woeif pwe fldskf
' powe woeif pwe fldskf