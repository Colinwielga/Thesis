VERSION 5.00
Begin VB.Form frmFavorites 
   BackColor       =   &H00FFC0FF&
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAsk 
      Caption         =   "Click to Find out a Favorite"
      Height          =   975
      Left            =   2760
      TabIndex        =   4
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton cmdOptions 
      Caption         =   "Click Here to See Options"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Back to Menu"
      Height          =   615
      Left            =   2760
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
   End
   Begin VB.PictureBox picTable 
      Height          =   1695
      Left            =   240
      ScaleHeight     =   1635
      ScaleWidth      =   1995
      TabIndex        =   1
      Top             =   720
      Width           =   2055
   End
   Begin VB.PictureBox picResults 
      Height          =   495
      Left            =   480
      ScaleHeight     =   435
      ScaleWidth      =   3555
      TabIndex        =   0
      Top             =   2520
      Width           =   3615
   End
End
Attribute VB_Name = "frmFavorites"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAsk_Click()
Dim Favorite As String
Dim Actor As String
Dim Actress As String
Dim Song As String
Dim Band As String
Dim Movie As String
Dim Book As String
Dim TVShow As String
Favorite = InputBox("Ask a Favorite")
picResults.Cls
If Favorite = "Actor" Then
    picResults.Print "Kristie's favorite actor is Adam Brody. "
End If
If Favorite = "Actress" Then
    picResults.Print "Kristie's favorite actress is Katie Holmes."
End If
If Favorite = "Song" Then
    picResults.Print "Kristie's favorite song is 'So Far Away.'"
End If
If Favorite = "Band" Then
    picResults.Print "Kristie's favorite band is Staind."
End If
If Favorite = "Movie" Then
    picResults.Print "Kristie's favorite movie is 'Step Up.'"
End If
If Favorite = "Book" Then
    picResults.Print "Kristie's favorite book is the Bible."
End If
If Favorite = "TVShow" Then
    picResults.Print "Kristie's favorite TV Show is 'Grey's Anatomy.'"
End If
End Sub

Private Sub cmdMenu_Click()
frmFavorites.Hide
frmMenu.Show
End Sub
Private Sub cmdOptions_Click()
picTable.Print "See Kristie's Favorite..."
picTable.Print "Actor"
picTable.Print "Actress"
picTable.Print "Song"
picTable.Print "Band"
picTable.Print "Movie"
picTable.Print "Book"
picTable.Print "TV Show"
End Sub

