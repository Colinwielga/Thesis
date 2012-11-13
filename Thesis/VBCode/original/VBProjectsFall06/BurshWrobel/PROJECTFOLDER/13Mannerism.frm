VERSION 5.00
Begin VB.Form Form14 
   Caption         =   "Form14"
   ClientHeight    =   8835
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11085
   LinkTopic       =   "Form14"
   Picture         =   "13Mannerism.frx":0000
   ScaleHeight     =   8835
   ScaleWidth      =   11085
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton Command13 
      Caption         =   "Fav"
      Height          =   375
      Left            =   1200
      TabIndex        =   23
      Top             =   8280
      Width           =   495
   End
   Begin VB.PictureBox Picture9 
      Height          =   1575
      Left            =   2640
      ScaleHeight     =   1515
      ScaleWidth      =   5955
      TabIndex        =   22
      Top             =   2760
      Visible         =   0   'False
      Width           =   6015
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Next ->"
      Height          =   615
      Left            =   9840
      TabIndex        =   21
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton Command11 
      Caption         =   "<- Previous"
      Height          =   615
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Introduction to Era"
      Height          =   375
      Left            =   4440
      TabIndex        =   19
      Top             =   2280
      Width           =   2295
   End
   Begin VB.PictureBox Picture8 
      Height          =   3255
      Left            =   9240
      ScaleHeight     =   3195
      ScaleWidth      =   1635
      TabIndex        =   18
      Top             =   4920
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.PictureBox Picture6 
      Height          =   3375
      Left            =   2400
      ScaleHeight     =   3315
      ScaleWidth      =   2955
      TabIndex        =   17
      Top             =   5040
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.PictureBox Picture7 
      Height          =   2055
      Left            =   5520
      ScaleHeight     =   1995
      ScaleWidth      =   3555
      TabIndex        =   16
      Top             =   5520
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.PictureBox Picture5 
      Height          =   3375
      Left            =   120
      ScaleHeight     =   3315
      ScaleWidth      =   2115
      TabIndex        =   15
      Top             =   2880
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton Command9 
      Caption         =   "C"
      Height          =   255
      Left            =   10680
      TabIndex        =   14
      Top             =   8280
      Width           =   375
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Info"
      Height          =   255
      Left            =   8880
      TabIndex        =   13
      Top             =   8280
      Width           =   1695
   End
   Begin VB.CommandButton Command7 
      Caption         =   "C"
      Height          =   255
      Left            =   7920
      TabIndex        =   12
      Top             =   7680
      Width           =   375
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Info"
      Height          =   255
      Left            =   6120
      TabIndex        =   11
      Top             =   7680
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "C"
      Height          =   255
      Left            =   4560
      TabIndex        =   10
      Top             =   8520
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Info"
      Height          =   255
      Left            =   2760
      TabIndex        =   9
      Top             =   8520
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "C"
      Height          =   255
      Left            =   1920
      TabIndex        =   8
      Top             =   6360
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Info"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   6360
      Width           =   1695
   End
   Begin VB.PictureBox Picture4 
      Height          =   3255
      Left            =   9240
      Picture         =   "13Mannerism.frx":D830
      ScaleHeight     =   3195
      ScaleWidth      =   1635
      TabIndex        =   6
      Top             =   4920
      Width           =   1695
   End
   Begin VB.PictureBox Picture3 
      Height          =   2055
      Left            =   5520
      Picture         =   "13Mannerism.frx":FFB0
      ScaleHeight     =   1995
      ScaleWidth      =   3555
      TabIndex        =   5
      Top             =   5520
      Width           =   3615
   End
   Begin VB.PictureBox Picture2 
      Height          =   3375
      Left            =   2400
      Picture         =   "13Mannerism.frx":13396
      ScaleHeight     =   3315
      ScaleWidth      =   2955
      TabIndex        =   4
      Top             =   5040
      Width           =   3015
   End
   Begin VB.PictureBox Picture1 
      Height          =   3375
      Left            =   120
      Picture         =   "13Mannerism.frx":16DDF
      ScaleHeight     =   3315
      ScaleWidth      =   2115
      TabIndex        =   3
      Top             =   2880
      Width           =   2175
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
      Caption         =   "1550 AD ~ 1600 AD"
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
      Height          =   1095
      Left            =   0
      TabIndex        =   1
      Top             =   1560
      Width           =   11055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Mannerism"
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
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   11055
   End
End
Attribute VB_Name = "Form14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'WesternArt Project
'Form14
'Bursh,Wrobel
'11-1-06
'This is our Mannerism Form Era, displaying works from Era.
Option Explicit
Private Sub Command1_Click()
    Form1.Show
    Form14.Hide
End Sub

Private Sub Command10_Click()
    Picture9.Visible = True
    Picture9.Cls
    Picture9.Print "    Mannerism - meaning an elegant, style refinement, was a result of social"
    Picture9.Print "and military disharmony, and the splitting of the Christian church into two camps."
    Picture9.Print "Emphasis was placed on God's role as judge, and focus on a pessimistic way of life."
    Picture9.Print "The main subject for Mannerist artists was the human body which was elongated,"
    Picture9.Print "exaggerated, elegant and/or arranged in a complex, erotic form which appealed to"
    Picture9.Print "a very elite audience."
End Sub

Private Sub Command11_Click()
Form13.Show
Form14.Hide
End Sub

Private Sub Command12_Click()
Form15.Show
Form14.Hide
End Sub

Private Sub Command13_Click()
    Form26.Show
    Form14.Hide
End Sub

Private Sub Command2_Click()
    Picture1.Visible = False
    Picture5.Visible = True
    Picture5.Cls
    Picture5.Print "Madonna of the Long Neck"
    Picture5.Print
    Picture5.Print "by Parmigianino"
    Picture5.Print
    Picture5.Print "1551 AD"
End Sub

Private Sub Command3_Click()
    Picture1.Visible = True
    Picture5.Visible = False
    Picture5.Cls
End Sub

Private Sub Command4_Click()
    Picture2.Visible = False
    Picture6.Visible = True
    Picture6.Cls
    Picture6.Print "Venus, Cupid, Folly and Time"
    Picture6.Print
    Picture6.Print "by Agnolo Bronzino"
    Picture6.Print
    Picture6.Print "1551 AD"
End Sub

Private Sub Command5_Click()
    Picture2.Visible = True
    Picture6.Visible = False
    Picture6.Cls
End Sub

Private Sub Command6_Click()
    Picture3.Visible = False
    Picture7.Visible = True
    Picture7.Cls
    Picture7.Print "The Last Supper"
    Picture7.Print
    Picture7.Print "by Jacopo Tintoretto"
    Picture7.Print
    Picture7.Print "1590 AD"
End Sub

Private Sub Command7_Click()
    Picture3.Visible = True
    Picture7.Visible = False
    Picture7.Cls
End Sub

Private Sub Command8_Click()
    Picture4.Visible = False
    Picture8.Visible = True
    Picture8.Cls
    Picture8.Print "The Resurrection"
    Picture8.Print "of Christ"
    Picture8.Print
    Picture8.Print "by El Greco"
    Picture8.Print
    Picture8.Print "1597 AD"
End Sub

Private Sub Command9_Click()
    Picture4.Visible = True
    Picture8.Visible = False
    Picture8.Cls
End Sub

Private Sub Picture9_Click()
    Picture9.Visible = False
End Sub
