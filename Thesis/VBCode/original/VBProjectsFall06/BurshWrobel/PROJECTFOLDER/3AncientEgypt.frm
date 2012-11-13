VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00000000&
   Caption         =   "Form4"
   ClientHeight    =   8835
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11085
   LinkTopic       =   "Form4"
   Picture         =   "3AncientEgypt.frx":0000
   ScaleHeight     =   8835
   ScaleWidth      =   11085
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command11 
      Caption         =   "Fav"
      Height          =   375
      Left            =   1200
      TabIndex        =   19
      Top             =   8280
      Width           =   495
   End
   Begin VB.PictureBox Picture7 
      Height          =   1215
      Left            =   1320
      ScaleHeight     =   1155
      ScaleWidth      =   8475
      TabIndex        =   18
      Top             =   3000
      Visible         =   0   'False
      Width           =   8535
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Next ->"
      Height          =   615
      Left            =   9840
      TabIndex        =   17
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      Caption         =   "<- Previous"
      Height          =   615
      Left            =   0
      TabIndex        =   16
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Introduction to Era"
      Height          =   375
      Left            =   4320
      TabIndex        =   15
      Top             =   2520
      Width           =   2295
   End
   Begin VB.CommandButton Command7 
      Caption         =   "C"
      Height          =   255
      Left            =   9720
      TabIndex        =   14
      Top             =   7320
      Width           =   375
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Info"
      Height          =   255
      Left            =   7920
      TabIndex        =   13
      Top             =   7320
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "C"
      Height          =   255
      Left            =   6360
      TabIndex        =   12
      Top             =   7320
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Info"
      Height          =   255
      Left            =   4560
      TabIndex        =   11
      Top             =   7320
      Width           =   1695
   End
   Begin VB.PictureBox Picture6 
      Height          =   2775
      Left            =   7800
      ScaleHeight     =   2715
      ScaleWidth      =   2355
      TabIndex        =   10
      Top             =   4440
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.PictureBox Picture5 
      Height          =   2775
      Left            =   4440
      ScaleHeight     =   2715
      ScaleWidth      =   2235
      TabIndex        =   9
      Top             =   4440
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.PictureBox Picture4 
      Height          =   2295
      Left            =   720
      ScaleHeight     =   2235
      ScaleWidth      =   2715
      TabIndex        =   8
      Top             =   4560
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton Command3 
      Caption         =   "C"
      Height          =   255
      Left            =   2760
      TabIndex        =   7
      Top             =   6960
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Info"
      Height          =   255
      Left            =   960
      TabIndex        =   6
      Top             =   6960
      Width           =   1695
   End
   Begin VB.PictureBox Picture3 
      Height          =   2775
      Left            =   7800
      Picture         =   "3AncientEgypt.frx":EE79
      ScaleHeight     =   2715
      ScaleWidth      =   2355
      TabIndex        =   5
      Top             =   4440
      Width           =   2415
   End
   Begin VB.PictureBox Picture2 
      Height          =   2775
      Left            =   4440
      Picture         =   "3AncientEgypt.frx":1058B
      ScaleHeight     =   2715
      ScaleWidth      =   2235
      TabIndex        =   4
      Top             =   4440
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      Height          =   2295
      Left            =   720
      Picture         =   "3AncientEgypt.frx":11F03
      ScaleHeight     =   2235
      ScaleWidth      =   2715
      TabIndex        =   3
      Top             =   4560
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Main Menu"
      Height          =   500
      Left            =   120
      TabIndex        =   2
      Top             =   8160
      Width           =   1000
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "2,500 BCE ~ 450 BCE"
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
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   10935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ancient Egypt"
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
      Width           =   10935
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'WesternArt Project
'Form4
'Bursh,Wrobel
'11-1-06
'This is our Ancient Egypt Form Era, displaying works from Era.
Option Explicit
Private Sub Command1_Click()
    Form1.Show
    Form4.Hide
End Sub

Private Sub Command10_Click()
Form5.Show
Form4.Hide
End Sub

Private Sub Command11_Click()
Form26.Show
Form4.Hide
End Sub

Private Sub Command2_Click()
    Picture1.Visible = False
    Picture4.Visible = True
    Picture4.Cls
    Picture4.Print "Pyramids at Giza"
    Picture4.Print
    Picture4.Print "2,500 BCE"
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
    Picture5.Print "Coffin of Tutankhamo"
    Picture5.Print
    Picture5.Print "1,325 BCE"
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
    Picture6.Print "Temple of Ramses II"
    Picture6.Print
    Picture6.Print "1,250 BCE"
End Sub

Private Sub Command7_Click()
    Picture3.Visible = True
    Picture6.Visible = False
    Picture6.Cls
End Sub


Private Sub Command8_Click()
    Picture7.Visible = True
    Picture7.Cls
    Picture7.Print "    'The Gift of the Nile' drew the attention of farmers before 7,000 BCE, but it wasn't until roughly 2,700 BCE"
    Picture7.Print "that Egypt united and became a kingdom under a Pharaoh.  The Pharaoh, considered a God on Earth, was inpregnated"
    Picture7.Print "each reign by Ra to bore a son.  The Egyptians were polytheists, who believed in Gods who represented all likes"
    Picture7.Print "of nature.  To preserve life after death they were physically survived by possesions that reminded them of their"
    Picture7.Print "daily activities.  Pyramids housed the tombs of many Pharaohs."
End Sub

Private Sub Command9_Click()
Form3.Show
Form4.Hide
End Sub

Private Sub Picture7_Click()
    Picture7.Visible = False
End Sub
