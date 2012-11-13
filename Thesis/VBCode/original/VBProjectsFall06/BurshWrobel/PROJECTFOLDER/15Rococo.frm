VERSION 5.00
Begin VB.Form Form16 
   Caption         =   "Form16"
   ClientHeight    =   8835
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11085
   LinkTopic       =   "Form16"
   Picture         =   "15Rococo.frx":0000
   ScaleHeight     =   8835
   ScaleWidth      =   11085
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command11 
      Caption         =   "Fav"
      Height          =   375
      Left            =   1200
      TabIndex        =   19
      Top             =   8400
      Width           =   495
   End
   Begin VB.PictureBox Picture7 
      Height          =   1095
      Left            =   2520
      ScaleHeight     =   1035
      ScaleWidth      =   6075
      TabIndex        =   18
      Top             =   3000
      Visible         =   0   'False
      Width           =   6135
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
      Left            =   4440
      TabIndex        =   15
      Top             =   2160
      Width           =   2295
   End
   Begin VB.PictureBox Picture6 
      Height          =   3255
      Left            =   8160
      ScaleHeight     =   3195
      ScaleWidth      =   2595
      TabIndex        =   14
      Top             =   4560
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.PictureBox Picture5 
      Height          =   2775
      Left            =   3480
      ScaleHeight     =   2715
      ScaleWidth      =   4155
      TabIndex        =   13
      Top             =   4800
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.PictureBox Picture4 
      Height          =   3255
      Left            =   360
      ScaleHeight     =   3195
      ScaleWidth      =   2595
      TabIndex        =   12
      Top             =   4560
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton Command7 
      Caption         =   "C"
      Height          =   255
      Left            =   10200
      TabIndex        =   11
      Top             =   7920
      Width           =   375
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Info"
      Height          =   255
      Left            =   8400
      TabIndex        =   10
      Top             =   7920
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "C"
      Height          =   255
      Left            =   6120
      TabIndex        =   9
      Top             =   7680
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Info"
      Height          =   255
      Left            =   4320
      TabIndex        =   8
      Top             =   7680
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "C"
      Height          =   255
      Left            =   2280
      TabIndex        =   7
      Top             =   7920
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Info"
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   7920
      Width           =   1695
   End
   Begin VB.PictureBox Picture3 
      Height          =   3255
      Left            =   8160
      Picture         =   "15Rococo.frx":294E9
      ScaleHeight     =   3195
      ScaleWidth      =   2595
      TabIndex        =   5
      Top             =   4560
      Width           =   2655
   End
   Begin VB.PictureBox Picture2 
      Height          =   2775
      Left            =   3480
      Picture         =   "15Rococo.frx":2BEC1
      ScaleHeight     =   2715
      ScaleWidth      =   4155
      TabIndex        =   4
      Top             =   4800
      Width           =   4215
   End
   Begin VB.PictureBox Picture1 
      Height          =   3255
      Left            =   360
      Picture         =   "15Rococo.frx":2E8FA
      ScaleHeight     =   3195
      ScaleWidth      =   2595
      TabIndex        =   3
      Top             =   4560
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
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
      Caption         =   "1710 AD ~ 1750 AD"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   23.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   1
      Top             =   1440
      Width           =   11055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Rococo"
      BeginProperty Font 
         Name            =   "High Tower Text"
         Size            =   69
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11055
   End
End
Attribute VB_Name = "Form16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'WesternArt Project
'Form16
'Bursh,Wrobel
'11-1-06
'This is our Rococo Form Era, displaying works from Era.
Option Explicit
Private Sub Command1_Click()
    Form1.Show
    Form16.Hide
End Sub

Private Sub Command10_Click()
Form17.Show
Form16.Hide
End Sub

Private Sub Command11_Click()
Form26.Show
Form16.Hide
End Sub

Private Sub Command2_Click()
    Picture1.Visible = False
    Picture4.Visible = True
    Picture4.Cls
    Picture4.Print "The Zwinger"
    Picture4.Print
    Picture4.Print "by Matthäus Daniel Pöppelmann"
    Picture4.Print
    Picture4.Print "1715 AD"
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
    Picture5.Print "Pilgrimage to Cythera"
    Picture5.Print
    Picture5.Print "by Antoine Watteau"
    Picture5.Print
    Picture5.Print "1717 AD"
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
    Picture6.Print "The Swing"
    Picture6.Print
    Picture6.Print "by Jean-Honoré Fragonard"
    Picture6.Print
    Picture6.Print "1750 AD"
End Sub

Private Sub Command7_Click()
    Picture3.Visible = True
    Picture6.Visible = False
    Picture6.Cls
End Sub

Private Sub Command8_Click()
    Picture7.Visible = True
    Picture7.Cls
    Picture7.Print "    Rococo meaning literally - rock- and -shell- The Age of Enlightenment brought forth"
    Picture7.Print "the Rococo style, which is an expression of wit and frivolity with somber, satrical"
    Picture7.Print "undertakings in a world of fantasy and grace."

End Sub

Private Sub Command9_Click()
Form15.Show
Form16.Hide
End Sub

Private Sub Picture7_Click()
Picture7.Visible = False
End Sub
