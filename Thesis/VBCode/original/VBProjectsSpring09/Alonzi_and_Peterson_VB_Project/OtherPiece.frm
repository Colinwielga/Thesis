VERSION 5.00
Begin VB.Form OtherPiece 
   BackColor       =   &H0000FFFF&
   Caption         =   "What other pieces are in the key of the first note you chose?"
   ClientHeight    =   7770
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8820
   LinkTopic       =   "Form1"
   ScaleHeight     =   7770
   ScaleWidth      =   8820
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Artistss 
      BackColor       =   &H80000002&
      Caption         =   "Let me see some of the artists, please!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6840
      Width           =   8415
   End
   Begin VB.CommandButton ShowAll 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Show me all the pieces in alphabetical order, please."
      Height          =   975
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5760
      Width           =   8415
   End
   Begin VB.CommandButton GoBack 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Let's play more Piano!"
      Height          =   1095
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4440
      Width           =   2895
   End
   Begin VB.CommandButton Both 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Show me both Major and Minor, please."
      Height          =   1095
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3120
      Width           =   2895
   End
   Begin VB.CommandButton Minor 
      BackColor       =   &H00FF80FF&
      Caption         =   "I want Minor Songs"
      Height          =   1215
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1680
      Width           =   2895
   End
   Begin VB.CommandButton Major 
      BackColor       =   &H0080FF80&
      Caption         =   "What are the Major Songs?"
      Height          =   1215
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   2895
   End
   Begin VB.PictureBox KeyBox 
      BackColor       =   &H00FFFF80&
      Height          =   5175
      Left            =   3240
      ScaleHeight     =   5115
      ScaleWidth      =   5235
      TabIndex        =   0
      Top             =   360
      Width           =   5295
   End
End
Attribute VB_Name = "OtherPiece"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Artists_Click()
    
End Sub

Private Sub Artistss_Click()
    OtherPiece.Hide
    Artists.Show
End Sub

'Palonzison Piano
'This is the Piano Form
'Matthew Peterson and Nicholas Alonzi are the authors of this Form
'This form was written in 2009 in the month of March
'This form is written so that the user can look at songs that are in the key
    'of the note he or she decided to start on
'This page prints in a picture box, searches arrays and the last button sorts
    'the array and puts it in a text box on a different form which is
    'formatted to have a scroll bar.
Private Sub Both_Click()
Dim j As Integer
    KeyBox.Cls
    j = 1
    KeyBox.Print "Songs in "; Key; " Major and Minor are..."
    KeyBox.Print "*****************************************"
    For j = 1 To Ctr
        If KeySig(j) = Key & " major" Or KeySig(j) = Key & " minor" Then
            KeyBox.Print Piece(j)
        End If
    Next j
            
End Sub

Private Sub GoBack_Click()
    Piano.Show
    OtherPiece.Hide
End Sub



Private Sub Major_Click()
Dim j As Integer
    KeyBox.Cls
    j = 1
    KeyBox.Print "Songs in "; Key; " Major are..."
    KeyBox.Print "*****************************************"
    For j = 1 To Ctr
        If KeySig(j) = Key & " major" Then
            KeyBox.Print Piece(j)
        End If
    Next j
            
End Sub

Private Sub Minor_Click()
Dim j As Integer
    KeyBox.Cls
    j = 1
    KeyBox.Print "Songs in "; Key; " Minor are..."
    KeyBox.Print "*****************************************"
    For j = 1 To Ctr
        If KeySig(j) = Key & " minor" Then
            KeyBox.Print Piece(j)
        End If
    Next j
            
End Sub

Private Sub ShowAll_Click()
Dim pass As Integer, pos As Integer, j As Integer
Dim tempSong As String, tempKey As String, WholeList As String
j = 1
    For pass = 1 To Ctr - 1
        For pos = 1 To Ctr - pass
            If Piece(pos) < Piece(pos + 1) Then
                tempSong = Piece(pos)
                Piece(pos) = Piece(pos + 1)
                Piece(pos + 1) = tempSong
                tempKey = KeySig(pos)
                KeySig(pos) = KeySig(pos + 1)
                KeySig(pos + 1) = tempKey
            End If
        Next pos
    Next pass
    AllSongs.Show
    For j = 1 To Ctr
        WholeList = Piece(j) & " ---> " & KeySig(j) & vbCrLf & WholeList
    Next j
    AllSongs.SongList.Text = WholeList
End Sub

