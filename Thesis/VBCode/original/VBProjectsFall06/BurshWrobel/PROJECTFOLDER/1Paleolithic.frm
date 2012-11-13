VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   8835
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11085
   LinkTopic       =   "Form2"
   ScaleHeight     =   8835
   ScaleWidth      =   11085
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   8895
      Left            =   0
      Picture         =   "1Paleolithic.frx":0000
      ScaleHeight     =   8835
      ScaleWidth      =   10995
      TabIndex        =   0
      Top             =   0
      Width           =   11055
      Begin VB.CommandButton Command8 
         Caption         =   "Fav"
         Height          =   375
         Left            =   1200
         TabIndex        =   15
         Top             =   8400
         Width           =   495
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Next ->"
         Height          =   615
         Left            =   9840
         TabIndex        =   14
         Top             =   120
         Width           =   1215
      End
      Begin VB.PictureBox Picture6 
         Height          =   1215
         Left            =   1920
         ScaleHeight     =   1155
         ScaleWidth      =   7275
         TabIndex        =   13
         Top             =   3240
         Visible         =   0   'False
         Width           =   7335
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Introduction to Era"
         Height          =   375
         Left            =   4320
         TabIndex        =   12
         Top             =   2760
         Width           =   2295
      End
      Begin VB.PictureBox Picture5 
         Height          =   2415
         Left            =   6000
         ScaleHeight     =   2355
         ScaleWidth      =   3435
         TabIndex        =   11
         Top             =   5280
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.CommandButton Command5 
         Caption         =   "C"
         Height          =   255
         Left            =   8040
         TabIndex        =   10
         Top             =   7800
         Width           =   375
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Info"
         Height          =   255
         Left            =   6360
         TabIndex        =   9
         Top             =   7800
         Width           =   1575
      End
      Begin VB.PictureBox Picture4 
         FontTransparent =   0   'False
         Height          =   3255
         Left            =   1680
         ScaleHeight     =   3195
         ScaleWidth      =   2715
         TabIndex        =   8
         Top             =   4560
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.CommandButton Command3 
         Caption         =   "C"
         Height          =   255
         Left            =   3720
         TabIndex        =   7
         Top             =   7920
         Width           =   375
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Info"
         Height          =   255
         Left            =   1920
         TabIndex        =   6
         Top             =   7920
         Width           =   1695
      End
      Begin VB.PictureBox Picture3 
         Height          =   2415
         Left            =   6000
         Picture         =   "1Paleolithic.frx":15F6F
         ScaleHeight     =   2355
         ScaleWidth      =   3435
         TabIndex        =   5
         Top             =   5280
         Width           =   3495
      End
      Begin VB.PictureBox Picture2 
         Height          =   3255
         Left            =   1680
         Picture         =   "1Paleolithic.frx":19FFB
         ScaleHeight     =   3195
         ScaleWidth      =   2715
         TabIndex        =   4
         Top             =   4560
         Width           =   2775
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H8000000A&
         Caption         =   "Main Menu"
         Height          =   500
         Left            =   120
         TabIndex        =   3
         Top             =   8280
         Width           =   1000
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "25,000 BCE ~ 6,000 BCE"
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
         Left            =   0
         TabIndex        =   2
         Top             =   2040
         Width           =   11055
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Paleolithic "
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
         Height          =   1335
         Left            =   0
         TabIndex        =   1
         Top             =   600
         Width           =   11055
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'WesternArt Project
'Form2
'Bursh,Wrobel
'11-1-06
'This is our Paleolithic Form Era, displaying works from Era.
Option Explicit
Private Sub Command1_Click()
    Form1.Show
    Form2.Hide
End Sub

Private Sub Command2_Click()
    Picture2.Visible = False
    Picture4.Visible = True
    Picture4.Cls
    Picture4.Print "Venus of Willendorf"
    Picture4.Print
    Picture4.Print "23,000 BCE"
End Sub
Private Sub Command3_Click()
    Picture2.Visible = True
    Picture4.Visible = False
    Picture4.Cls
End Sub

Private Sub Command4_Click()
    Picture3.Visible = False
    Picture5.Visible = True
    Picture5.Cls
    Picture5.Print "Chinese Horse"
    Picture5.Print
    Picture5.Print "14,000 BCE"
End Sub

Private Sub Command5_Click()
    Picture3.Visible = True
    Picture5.Visible = False
    Picture5.Cls
End Sub

Private Sub Command6_Click()
    Picture6.Visible = True
    Picture6.Cls
    Picture6.Print "    By 50,000 BCE Homo sapiens sapiens - wise men - had begun developing complex cultures with"
    Picture6.Print "the daily life of hunting and gathering.  They built dwellings within cave areas made from animal skins,  "
    Picture6.Print "mud and stone.  There is evidence of language through marks found on hard surfaces, for symbolic"
    Picture6.Print "purposes."
End Sub

Private Sub Command7_Click()
    Form3.Show
    Form2.Hide
End Sub


Private Sub Command8_Click()
    Form26.Show
    Form2.Hide
End Sub

Private Sub Picture6_Click()
    Picture6.Visible = False
End Sub
