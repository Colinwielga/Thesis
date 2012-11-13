VERSION 5.00
Begin VB.Form frmInfo 
   Caption         =   "Information on Weezer's band members"
   ClientHeight    =   8625
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10860
   LinkTopic       =   "Form1"
   Picture         =   "frmInfo.frx":0000
   ScaleHeight     =   8625
   ScaleWidth      =   10860
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Current Data"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9000
      TabIndex        =   10
      Top             =   3720
      Width           =   1695
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Exit this rad program"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   8
      Top             =   7440
      Width           =   3855
   End
   Begin VB.CommandButton cmdSchedule 
      Caption         =   "Go to Weezer's Tour Schedule"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   7
      Top             =   6000
      Width           =   3855
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Go back to the main page"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   6
      Top             =   6720
      Width           =   3855
   End
   Begin VB.CommandButton cmdPictures 
      Caption         =   "See Pictures of Weezer"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   5
      Top             =   5160
      Width           =   4095
   End
   Begin VB.CommandButton cmdPat 
      Caption         =   "Pat Wilson"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   4
      Top             =   4080
      Width           =   4095
   End
   Begin VB.CommandButton cmdScott 
      Caption         =   "Scott Shriner"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   3
      Top             =   3000
      Width           =   4095
   End
   Begin VB.CommandButton cmdBrian 
      Caption         =   "Brian Bell"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   2
      Top             =   1920
      Width           =   4095
   End
   Begin VB.CommandButton cmdRivers 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Rivers Cuomo"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      MaskColor       =   &H00400000&
      TabIndex        =   1
      Top             =   840
      Width           =   4095
   End
   Begin VB.PictureBox picResults 
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   4560
      ScaleHeight     =   3435
      ScaleWidth      =   6075
      TabIndex        =   0
      Top             =   120
      Width           =   6135
   End
   Begin VB.Label lblInfo 
      Caption         =   "Click on the band member's name to learn more about him"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   9
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Weezer
'Form Name: frmInfo.frm
'Author: Emily Balamut
'Date Written: 10/30/08
'Objective: This form allows the user to click on a button and see information on
'that particular band member.
Option Explicit

Private Sub cmdBack_Click()
    frmInfo.Hide
    frmBeginning.Show
End Sub

Private Sub cmdBrian_Click()
    picResults.Cls
    picResults.Print "Name: Brian Bell"
    picResults.Print "Position in Weezer: Rhythm guitarist, occasional vocalist"
    picResults.Print "Birthdate: December 9, 1968"
    picResults.Print "Birthplace: Iowa City, Iowa"
    picResults.Print "*******************************************************"
    picResults.Print "Brian started playing guitar when he was a freshman in"
    picResults.Print "high school. Instead of going to college, Brian moved to"
    picResults.Print "Los Angeles when he had finished high school, believing"
    picResults.Print "college to be a waste of money. He joined Weezer in 1993"
    picResults.Print "after Rivers and the old bassist, Matt Sharp, saw him"
    picResults.Print "perform in Los Angeles."
End Sub

Private Sub cmdClear_Click()
    picResults.Cls
End Sub

Private Sub cmdPat_Click()
    picResults.Cls
    picResults.Print "Name: Pat Wilson"
    picResults.Print "Position in Weezer: Drummer, Songwriter, Vocals"
    picResults.Print "Birthdate: February 1, 1969"
    picResults.Print "Birthplace: Buffalo, New York"
    picResults.Print "*******************************************************"
    picResults.Print "When he was 15, Pat began taking drum lessons after"
    picResults.Print "attending a Van Halen concert. Wilson dropped out of"
    picResults.Print "college after one semester, saying that college is"
    picResults.Print "'such bunk. Too much politics and jockeying for favor.'"
    picResults.Print "He moved to Los Angeles when he was 21 because he didn't"
    picResults.Print "like the local music scene. Besides Rivers, Pat is the"
    picResults.Print "only other original member of Weezer, playing with him"
    picResults.Print "since the beginning."
End Sub

Private Sub cmdPictures_Click()
    frmInfo.Hide
    frmPictures.Show
End Sub

Private Sub cmdQuit_Click()
MsgBox "Thanks for rocking out with Weezer, " & UserName & "! See you later!", , "Leave"
End
End Sub

Private Sub cmdRivers_Click()
    picResults.Cls
    picResults.Print "Name: Rivers Cuomo"
    picResults.Print "Position in Weezer: Lead Vocalist, songwriter, guitarist"
    picResults.Print "Birthdate: June 13, 1970"
    picResults.Print "Birthplace: New York City, New York"
    picResults.Print "*******************************************************"
    picResults.Print "Rivers grew up on an ashram. His family moved to "
    picResults.Print "Connecticut and this is where he first became interested"
    picResults.Print "in music and played in several bands. In 1989, he moved"
    picResults.Print "to Los Angeles to further his music career. He formed"
    picResults.Print "Weezer on Valentine's Day of 1992."

End Sub

Private Sub cmdSchedule_Click()
    frmInfo.Hide
    frmSchedule.Show
End Sub

Private Sub cmdScott_Click()
    picResults.Cls
    picResults.Print "Name: Scott Shriner"
    picResults.Print "Position in Weezer: Bass Guitar, Vocals"
    picResults.Print "Birthdate: July 11, 1965"
    picResults.Print "Birthplace: Toledo, Ohio"
    picResults.Print "*******************************************************"
    picResults.Print "Scott began playing bass guitar in high school. He joined"
    picResults.Print "the Marines out of high school, but soon decided to take"
    picResults.Print "up bass playing professionally. At the age of 25, he"
    picResults.Print "moved to Los Angeles and played with a number of bands."
    picResults.Print "In 2001, he joined Weezer as a back-up bassist, but soon"
    picResults.Print "became their full time bassist after the other one quit."
End Sub
