VERSION 5.00
Begin VB.Form Form10 
   BackColor       =   &H00FF8080&
   Caption         =   "Form10"
   ClientHeight    =   8445
   ClientLeft      =   60
   ClientTop       =   660
   ClientWidth     =   12165
   LinkTopic       =   "Form10"
   ScaleHeight     =   8445
   ScaleWidth      =   12165
   Begin VB.CommandButton cmdKey 
      Caption         =   "The Key to Programming Success"
      BeginProperty Font 
         Name            =   "PaintStroke"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3360
      TabIndex        =   5
      Top             =   6960
      Width           =   2295
   End
   Begin VB.CommandButton cmdBulletinBoards 
      Caption         =   "Pictures of Past Bulletin Boards"
      BeginProperty Font 
         Name            =   "PaintStroke"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   360
      TabIndex        =   3
      Top             =   6960
      Width           =   2295
   End
   Begin VB.PictureBox pbxResults 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "PaintStroke"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   120
      ScaleHeight     =   5955
      ScaleWidth      =   11835
      TabIndex        =   2
      Top             =   840
      Width           =   11895
   End
   Begin VB.CommandButton cmdProgramMenu 
      Caption         =   "Back to Programming"
      BeginProperty Font 
         Name            =   "PaintStroke"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   6360
      TabIndex        =   1
      Top             =   6960
      Width           =   2295
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Back to Main Menu"
      BeginProperty Font 
         Name            =   "PaintStroke"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   9360
      TabIndex        =   0
      Top             =   6960
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Caption         =   "Passive Programming"
      BeginProperty Font 
         Name            =   "PaintStroke"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   855
      Left            =   3840
      TabIndex        =   4
      Top             =   0
      Width           =   4815
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim M As String
Dim January As String, February As String, March As String, April As String
Dim May As String, June As String, July As String, August As String
Dim September As String, October As String, November As String, December As String

Private Sub cmdBulletinBoards_Click()
M = InputBox("Please Enter a Month", "Month")
pbxResults.Cls

If M = "January" Then
    pbxResults.Cls
    pbxResults.Picture = LoadPicture(strPath & "JanuaryPic.jpg")
ElseIf M = "February" Then
    pbxResults.Cls
    pbxResults.Print "Sorry. I have no picture from February. Please try another month."
ElseIf M = "March" Then
    pbxResults.Cls
    pbxResults.Print "Sorry. I have no pictures from March.  Please try another month."
ElseIf M = "April" Then
    pbxResults.Cls
    pbxResults.Print "Sorry. I have no picture from April. Please try another month."
ElseIf M = "May" Then
    pbxResults.Cls
    pbxResults.Print "Sorry. I have no picture from May. Please try another month."
ElseIf M = "June" Then
    pbxResults.Cls
    pbxResults.Print "Sorry. I have no picture from June. Please try another month."
ElseIf M = "July" Then
    pbxResults.Cls
    pbxResults.Print "Sorry. I have no picture from July. Please try another month."
ElseIf M = "August" Then
    pbxResults.Cls
    pbxResults.Print "Sorry. I have no picture from August. Please try another month."
ElseIf M = "September" Then
    pbxResults.Cls
    pbxResults.Picture = LoadPicture(strPath & "SeptemberPic.jpg")
ElseIf M = "October" Then
    pbxResults.Cls
    pbxResults.Print "Sorry. I have no picture from October. Please try another month."
ElseIf M = "November" Then
    pbxResults.Cls
    pbxResults.Picture = LoadPicture(strPath & "NovemberPic.jpg")
ElseIf M = "December" Then
    pbxResults.Cls
    pbxResults.Picture = LoadPicture(strPath & "DecemberPic.jpg")
End If
End Sub

Private Sub cmdKey_Click()
pbxResults.Print "Creativity!!!"
End Sub

Private Sub cmdMenu_Click()
Form10.Hide
Form2.Show
End Sub

Private Sub cmdProgramMenu_Click()
Form10.Hide
Form4.Show
End Sub
