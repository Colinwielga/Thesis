VERSION 5.00
Begin VB.Form Form13 
   Caption         =   "Form13"
   ClientHeight    =   8835
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11085
   LinkTopic       =   "Form13"
   Picture         =   "12HighRenaissance.frx":0000
   ScaleHeight     =   8835
   ScaleWidth      =   11085
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton Command17 
      Caption         =   "Fav"
      Height          =   375
      Left            =   1080
      TabIndex        =   31
      Top             =   8520
      Width           =   495
   End
   Begin VB.PictureBox Picture13 
      Height          =   2295
      Left            =   3840
      ScaleHeight     =   2235
      ScaleWidth      =   3195
      TabIndex        =   30
      Top             =   2640
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Next ->"
      Height          =   615
      Left            =   9840
      TabIndex        =   29
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton Command15 
      Caption         =   "<- Previous"
      Height          =   615
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Introduction to Era"
      Height          =   375
      Left            =   4320
      TabIndex        =   27
      Top             =   2160
      Width           =   2295
   End
   Begin VB.PictureBox Picture12 
      Height          =   2415
      Left            =   7920
      ScaleHeight     =   2355
      ScaleWidth      =   2955
      TabIndex        =   26
      Top             =   5760
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.PictureBox Picture11 
      Height          =   1695
      Left            =   7200
      ScaleHeight     =   1635
      ScaleWidth      =   3555
      TabIndex        =   25
      Top             =   3480
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.PictureBox Picture10 
      Height          =   3015
      Left            =   5400
      ScaleHeight     =   2955
      ScaleWidth      =   2115
      TabIndex        =   24
      Top             =   5400
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.PictureBox Picture9 
      Height          =   3015
      Left            =   3000
      ScaleHeight     =   2955
      ScaleWidth      =   2115
      TabIndex        =   23
      Top             =   5400
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.PictureBox Picture8 
      Height          =   2415
      Left            =   480
      ScaleHeight     =   2355
      ScaleWidth      =   1875
      TabIndex        =   22
      Top             =   5520
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.PictureBox Picture7 
      Height          =   1695
      Left            =   120
      ScaleHeight     =   1635
      ScaleWidth      =   3555
      TabIndex        =   21
      Top             =   3480
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.CommandButton Command13 
      Caption         =   "C"
      Height          =   255
      Left            =   9960
      TabIndex        =   20
      Top             =   8280
      Width           =   375
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Info"
      Height          =   255
      Left            =   8160
      TabIndex        =   19
      Top             =   8280
      Width           =   1695
   End
   Begin VB.CommandButton Command11 
      Caption         =   "C"
      Height          =   255
      Left            =   9600
      TabIndex        =   18
      Top             =   5280
      Width           =   375
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Info"
      Height          =   255
      Left            =   7800
      TabIndex        =   17
      Top             =   5280
      Width           =   1695
   End
   Begin VB.CommandButton Command9 
      Caption         =   "C"
      Height          =   255
      Left            =   7200
      TabIndex        =   16
      Top             =   8520
      Width           =   375
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Info"
      Height          =   255
      Left            =   5400
      TabIndex        =   15
      Top             =   8520
      Width           =   1695
   End
   Begin VB.CommandButton Command7 
      Caption         =   "C"
      Height          =   255
      Left            =   4800
      TabIndex        =   14
      Top             =   8520
      Width           =   375
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Info"
      Height          =   255
      Left            =   3000
      TabIndex        =   13
      Top             =   8520
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "C"
      Height          =   255
      Left            =   2280
      TabIndex        =   12
      Top             =   8040
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Info"
      Height          =   255
      Left            =   480
      TabIndex        =   11
      Top             =   8040
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "C"
      Height          =   255
      Left            =   2280
      TabIndex        =   10
      Top             =   5280
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Info"
      Height          =   255
      Left            =   480
      TabIndex        =   9
      Top             =   5280
      Width           =   1695
   End
   Begin VB.PictureBox Picture6 
      Height          =   2415
      Left            =   7920
      Picture         =   "12HighRenaissance.frx":1F19F
      ScaleHeight     =   2355
      ScaleWidth      =   2955
      TabIndex        =   8
      Top             =   5760
      Width           =   3015
   End
   Begin VB.PictureBox Picture5 
      Height          =   1695
      Left            =   7200
      Picture         =   "12HighRenaissance.frx":21FD0
      ScaleHeight     =   1635
      ScaleWidth      =   3555
      TabIndex        =   7
      Top             =   3480
      Width           =   3615
   End
   Begin VB.PictureBox Picture4 
      Height          =   3015
      Left            =   5400
      Picture         =   "12HighRenaissance.frx":2620F
      ScaleHeight     =   2955
      ScaleWidth      =   2115
      TabIndex        =   6
      Top             =   5400
      Width           =   2175
   End
   Begin VB.PictureBox Picture3 
      Height          =   3015
      Left            =   3000
      Picture         =   "12HighRenaissance.frx":279DC
      ScaleHeight     =   2955
      ScaleWidth      =   2115
      TabIndex        =   5
      Top             =   5400
      Width           =   2175
   End
   Begin VB.PictureBox Picture2 
      Height          =   2415
      Left            =   480
      Picture         =   "12HighRenaissance.frx":28F68
      ScaleHeight     =   2355
      ScaleWidth      =   1875
      TabIndex        =   4
      Top             =   5520
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      Height          =   1695
      Left            =   120
      Picture         =   "12HighRenaissance.frx":2AD91
      ScaleHeight     =   1635
      ScaleWidth      =   3555
      TabIndex        =   3
      Top             =   3480
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Main Menu"
      Height          =   500
      Left            =   0
      TabIndex        =   2
      Top             =   8400
      Width           =   1000
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1450 AD ~ 1550 AD"
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
      Height          =   855
      Left            =   -120
      TabIndex        =   1
      Top             =   1440
      Width           =   11055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "High Renaissance"
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
      Height          =   1575
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   11055
   End
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'WesternArt Project
'Form
'Bursh,Wrobel
'11-1-06
'This is our High Renaissance Form Era, displaying works from Era.
Option Explicit
Private Sub Command1_Click()
    Form1.Show
    Form13.Hide
End Sub

Private Sub Command10_Click()
    Picture5.Visible = False
    Picture11.Visible = True
    Picture11.Cls
    Picture11.Print "The Sistine Chapel Ceiling"
    Picture11.Print
    Picture11.Print "by Michaelangelo"
    Picture11.Print
    Picture11.Print "1510 AD"
End Sub

Private Sub Command11_Click()
    Picture5.Visible = True
    Picture11.Visible = False
    Picture11.Cls
End Sub

Private Sub Command12_Click()
    Picture6.Visible = False
    Picture12.Visible = True
    Picture12.Cls
    Picture12.Print "The School of Athlens"
    Picture12.Print
    Picture12.Print "by Raphael"
    Picture12.Print
    Picture12.Print "1510 AD"
End Sub

Private Sub Command13_Click()
    Picture6.Visible = True
    Picture12.Visible = False
    Picture12.Cls
End Sub

Private Sub Command14_Click()
    Picture13.Visible = True
    Picture13.Cls
    Picture13.Print "    With -high- reflecting upon the nature of"
    Picture13.Print " this Renaissance era, the High Renaissance"
    Picture13.Print "housed great accomplishments by a few "
    Picture13.Print "competing artistic personalities. Competition"
    Picture13.Print "within architecture, sculpture and painting "
    Picture13.Print "were however overlooked by politics, but"
    Picture13.Print "these works came to be some of the best the"
    Picture13.Print "Western World has ever seen."
End Sub

Private Sub Command15_Click()
Form12.Show
Form13.Hide
End Sub

Private Sub Command16_Click()
Form14.Show
Form13.Hide
End Sub

Private Sub Command17_Click()
    Form26.Show
    Form13.Hide
End Sub

Private Sub Command2_Click()
    Picture1.Visible = False
    Picture7.Visible = True
    Picture7.Cls
    Picture7.Print "The Last Supper"
    Picture7.Print
    Picture7.Print "by Leonardo da Vinci"
    Picture7.Print
    Picture7.Print "1496 AD"
End Sub

Private Sub Command3_Click()
    Picture1.Visible = True
    Picture7.Visible = False
    Picture7.Cls
End Sub

Private Sub Command4_Click()
    Picture2.Visible = False
    Picture8.Visible = True
    Picture8.Cls
    Picture8.Print "Piéta"
    Picture8.Print
    Picture8.Print "by Michaelangelo"
    Picture8.Print
    Picture8.Print "1499 AD"
End Sub

Private Sub Command5_Click()
    Picture2.Visible = True
    Picture8.Visible = False
    Picture8.Cls
End Sub

Private Sub Command6_Click()
    Picture3.Visible = False
    Picture9.Visible = True
    Picture9.Cls
    Picture9.Print "The Mona Lisa"
    Picture9.Print
    Picture9.Print "by Leonardo da Vinci"
    Picture9.Print
    Picture9.Print "1504 AD"
End Sub

Private Sub Command7_Click()
    Picture3.Visible = True
    Picture9.Visible = False
    Picture9.Cls
End Sub

Private Sub Command8_Click()
    Picture4.Visible = False
    Picture10.Visible = True
    Picture10.Cls
    Picture10.Print "David"
    Picture10.Print
    Picture10.Print "by Michaelangelo"
    Picture10.Print
    Picture10.Print "1504 AD"
End Sub

Private Sub Command9_Click()
    Picture4.Visible = True
    Picture10.Visible = False
    Picture10.Cls
End Sub

Private Sub Picture13_Click()
Picture13.Visible = False
End Sub
