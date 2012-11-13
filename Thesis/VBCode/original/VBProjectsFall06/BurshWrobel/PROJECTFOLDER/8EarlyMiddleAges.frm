VERSION 5.00
Begin VB.Form Form9 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Form9"
   ClientHeight    =   8835
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11085
   LinkTopic       =   "Form9"
   Picture         =   "8EarlyMiddleAges.frx":0000
   ScaleHeight     =   8835
   ScaleWidth      =   11085
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton Command11 
      Caption         =   "Fav"
      Height          =   375
      Left            =   1200
      TabIndex        =   19
      Top             =   8400
      Width           =   495
   End
   Begin VB.PictureBox Picture7 
      Height          =   1215
      Left            =   2040
      ScaleHeight     =   1155
      ScaleWidth      =   6555
      TabIndex        =   18
      Top             =   3240
      Visible         =   0   'False
      Width           =   6615
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Next ->"
      Height          =   615
      Left            =   9840
      TabIndex        =   17
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      Caption         =   "<- Previous"
      Height          =   615
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Introduction to Era"
      Height          =   375
      Left            =   4200
      TabIndex        =   15
      Top             =   2520
      Width           =   2295
   End
   Begin VB.CommandButton Command7 
      Caption         =   "C"
      Height          =   255
      Left            =   9840
      TabIndex        =   14
      Top             =   8400
      Width           =   375
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Info"
      Height          =   255
      Left            =   8040
      TabIndex        =   13
      Top             =   8400
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "C"
      Height          =   255
      Left            =   6120
      TabIndex        =   12
      Top             =   8160
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Info"
      Height          =   255
      Left            =   4320
      TabIndex        =   11
      Top             =   8160
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "C"
      Height          =   255
      Left            =   2520
      TabIndex        =   10
      Top             =   7920
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Info"
      Height          =   255
      Left            =   720
      TabIndex        =   9
      Top             =   7920
      Width           =   1695
   End
   Begin VB.PictureBox Picture6 
      Height          =   2895
      Left            =   7920
      ScaleHeight     =   2835
      ScaleWidth      =   2355
      TabIndex        =   8
      Top             =   5400
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.PictureBox Picture5 
      Height          =   1935
      Left            =   3840
      ScaleHeight     =   1875
      ScaleWidth      =   3315
      TabIndex        =   7
      Top             =   6120
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.PictureBox Picture4 
      Height          =   2895
      Left            =   600
      ScaleHeight     =   2835
      ScaleWidth      =   2355
      TabIndex        =   6
      Top             =   4920
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.PictureBox Picture3 
      Height          =   2895
      Left            =   7920
      Picture         =   "8EarlyMiddleAges.frx":10971
      ScaleHeight     =   2835
      ScaleWidth      =   2340
      TabIndex        =   5
      Top             =   5400
      Width           =   2400
   End
   Begin VB.PictureBox Picture2 
      Height          =   1935
      Left            =   3840
      Picture         =   "8EarlyMiddleAges.frx":12C34
      ScaleHeight     =   1875
      ScaleWidth      =   3315
      TabIndex        =   4
      Top             =   6120
      Width           =   3375
   End
   Begin VB.PictureBox Picture1 
      Height          =   2895
      Left            =   600
      Picture         =   "8EarlyMiddleAges.frx":15424
      ScaleHeight     =   2835
      ScaleWidth      =   2355
      TabIndex        =   3
      Top             =   4920
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000A&
      Caption         =   "Main Menu"
      Height          =   500
      Left            =   120
      TabIndex        =   2
      Top             =   8280
      Width           =   1000
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "850 AD ~ 1000 AD"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   23.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   1800
      Width           =   11055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Early Middle Ages"
      BeginProperty Font 
         Name            =   "High Tower Text"
         Size            =   69
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   11055
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'WesternArt Project
'Form9
'Bursh,Wrobel
'11-1-06
'This is our Early Middle Ages Form Era, displaying works from Era.
Option Explicit

Private Sub Command1_Click()
Form1.Show
Form9.Hide
End Sub

Private Sub Command10_Click()
Form10.Show
Form9.Hide
End Sub

Private Sub Command11_Click()
Form26.Show
Form9.Hide
End Sub

Private Sub Command2_Click()
    Picture1.Visible = False
    Picture4.Visible = True
    Picture4.Cls
    Picture4.Print "Christ in Majesty"
    Picture4.Print
    Picture4.Print "845 AD"
End Sub

Private Sub Command3_Click()
    Picture1.Visible = True
    Picture4.Visible = False
    Picture4.Cls
End Sub

Private Sub Command4_Click()
    Picture2.Visible = False
    Picture5.Visible = True
    Picture5.Cls
    Picture5.Print "The Great Mosque"
    Picture5.Print
    Picture5.Print "850 AD"
End Sub

Private Sub Command5_Click()
    Picture2.Visible = True
    Picture5.Visible = False
    Picture5.Cls
End Sub

Private Sub Command6_Click()
    Picture3.Visible = False
    Picture6.Visible = True
    Picture6.Cls
    Picture6.Print "The Four Evangelists"
    Picture6.Print
    Picture6.Print "850 AD"
End Sub

Private Sub Command7_Click()
    Picture3.Visible = True
    Picture6.Visible = False
    Picture6.Cls
End Sub

Private Sub Command8_Click()
    Picture7.Visible = True
    Picture7.Cls
    Picture7.Print "    With the Islamic breakthrough influencing Western art, manuscript illumination became"
    Picture7.Print "popular with the importance of the Bible in the spread of Christianity.  Monastaries were"
    Picture7.Print "built, where the religious ways of life in chasity, obedience, and poverty were practiced"
    Picture7.Print "daily."
End Sub

Private Sub Command9_Click()
Form8.Show
Form9.Hide
End Sub

Private Sub Picture7_Click()
Picture7.Visible = False
End Sub
