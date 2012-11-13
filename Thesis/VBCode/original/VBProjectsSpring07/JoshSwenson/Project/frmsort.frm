VERSION 5.00
Begin VB.Form frmsort 
   BackColor       =   &H00C000C0&
   Caption         =   "Countrys"
   ClientHeight    =   7755
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10590
   LinkTopic       =   "Form1"
   ScaleHeight     =   7755
   ScaleWidth      =   10590
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdlast 
      Caption         =   "Go to Last Page"
      Height          =   735
      Left            =   2040
      TabIndex        =   5
      Top             =   5280
      Width           =   1575
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "Go Back"
      Height          =   735
      Left            =   240
      TabIndex        =   4
      Top             =   5280
      Width           =   1575
   End
   Begin VB.TextBox txtcountry 
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   3255
   End
   Begin VB.CommandButton cmdend 
      Caption         =   "Quit"
      Height          =   735
      Left            =   360
      TabIndex        =   1
      Top             =   6600
      Width           =   1335
   End
   Begin VB.PictureBox picresults 
      BackColor       =   &H0000FFFF&
      Height          =   7095
      Left            =   3840
      ScaleHeight     =   7035
      ScaleWidth      =   6435
      TabIndex        =   0
      Top             =   360
      Width           =   6495
   End
   Begin VB.Label Label1 
      Caption         =   "Enter a country to see a list of players who have won more then 1 Grand Slam for that country."
      Height          =   615
      Left            =   360
      TabIndex        =   3
      Top             =   480
      Width           =   2775
   End
End
Attribute VB_Name = "frmsort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'navigate the pages
Private Sub cmdback_Click()
frmsort.Visible = False
frmpics.Visible = True
End Sub

Private Sub cmdend_Click()
End
End Sub



'navigate the pages
Private Sub cmdlast_Click()
frmsort.Visible = False
frmend.Visible = True

End Sub

Private Sub txtcountry_Change()
Dim player As String, pos As Integer, found As Integer
'as player inters a text the program automaticly searches for a result
player = txtcountry.Text
pos = 0
found = 0
picresults.Cls
picresults.Print "The following players from "; player; " have won atleast 2 Grand Slams"
picresults.Print "***************************************"
Do Until pos = ctr
pos = pos + 1
If country(pos) = player Then
picresults.Print names(pos); Tab(20); slams(pos)
found = found + 1
End If
Loop
If found = 0 Then
picresults.Print "No one from this country has won more then 1 Grand Slam."
End If
End Sub
