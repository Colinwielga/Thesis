VERSION 5.00
Begin VB.Form Form15 
   Caption         =   "Form15"
   ClientHeight    =   8835
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11085
   LinkTopic       =   "Form15"
   Picture         =   "14Baroque.frx":0000
   ScaleHeight     =   8835
   ScaleWidth      =   11085
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command13 
      Caption         =   "Fav"
      Height          =   375
      Left            =   1200
      TabIndex        =   23
      Top             =   8400
      Width           =   495
   End
   Begin VB.PictureBox Picture9 
      Height          =   1455
      Left            =   2640
      ScaleHeight     =   1395
      ScaleWidth      =   5715
      TabIndex        =   22
      Top             =   2880
      Visible         =   0   'False
      Width           =   5775
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Next - >"
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
      Top             =   2400
      Width           =   2295
   End
   Begin VB.PictureBox Picture8 
      Height          =   3015
      Left            =   8520
      ScaleHeight     =   2955
      ScaleWidth      =   2355
      TabIndex        =   18
      Top             =   4440
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.PictureBox Picture7 
      Height          =   3375
      Left            =   5640
      ScaleHeight     =   3315
      ScaleWidth      =   2595
      TabIndex        =   17
      Top             =   5040
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.PictureBox Picture6 
      Height          =   3375
      Left            =   2760
      ScaleHeight     =   3315
      ScaleWidth      =   2595
      TabIndex        =   16
      Top             =   5040
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.PictureBox Picture5 
      Height          =   3015
      Left            =   120
      ScaleHeight     =   2955
      ScaleWidth      =   2355
      TabIndex        =   15
      Top             =   4320
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Command9 
      Caption         =   "C"
      Height          =   255
      Left            =   10200
      TabIndex        =   14
      Top             =   7560
      Width           =   375
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Info"
      Height          =   255
      Left            =   8760
      TabIndex        =   13
      Top             =   7560
      Width           =   1335
   End
   Begin VB.CommandButton Command7 
      Caption         =   "C"
      Height          =   255
      Left            =   7680
      TabIndex        =   12
      Top             =   8520
      Width           =   375
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Info"
      Height          =   255
      Left            =   5880
      TabIndex        =   11
      Top             =   8520
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "C"
      Height          =   255
      Left            =   4800
      TabIndex        =   10
      Top             =   8520
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Info"
      Height          =   255
      Left            =   3000
      TabIndex        =   9
      Top             =   8520
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "C"
      Height          =   255
      Left            =   1920
      TabIndex        =   8
      Top             =   7440
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Info"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   7440
      Width           =   1695
   End
   Begin VB.PictureBox Picture4 
      Height          =   3015
      Left            =   8520
      Picture         =   "14Baroque.frx":C025
      ScaleHeight     =   2955
      ScaleWidth      =   2355
      TabIndex        =   6
      Top             =   4440
      Width           =   2415
   End
   Begin VB.PictureBox Picture3 
      Height          =   3375
      Left            =   5640
      Picture         =   "14Baroque.frx":DB08
      ScaleHeight     =   3315
      ScaleWidth      =   2595
      TabIndex        =   5
      Top             =   5040
      Width           =   2655
   End
   Begin VB.PictureBox Picture2 
      Height          =   3375
      Left            =   2760
      Picture         =   "14Baroque.frx":1044A
      ScaleHeight     =   3315
      ScaleWidth      =   2595
      TabIndex        =   4
      Top             =   5040
      Width           =   2655
   End
   Begin VB.PictureBox Picture1 
      Height          =   3015
      Left            =   120
      Picture         =   "14Baroque.frx":13031
      ScaleHeight     =   2955
      ScaleWidth      =   2355
      TabIndex        =   3
      Top             =   4320
      Width           =   2415
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
      Caption         =   "1600 AD ~ 1710 AD"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   23.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   1
      Top             =   1800
      Width           =   11055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Baroque"
      BeginProperty Font 
         Name            =   "High Tower Text"
         Size            =   69
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   11055
   End
End
Attribute VB_Name = "Form15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'WesternArt Project
'Form15
'Bursh,Wrobel
'11-1-06
'This is our Baroque Form Era, displaying works from Era.
Option Explicit
Private Sub Command1_Click()
    Form1.Show
    Form15.Hide
End Sub

Private Sub Command10_Click()
    Picture9.Visible = True
    Picture9.Cls
    Picture9.Print "    As conflicts between Catholics and Protestants continued, the Baroque"
    Picture9.Print "style became to be unrestrained, overtly emotional, and very energetic,"
    Picture9.Print "reflecting the many scientific advances of the period.  With Paris, France"
    Picture9.Print "becoming the new center  of the artistic world, the Baroque style contrasted"
    Picture9.Print "lights and darks, smooth and rough textures in many different asymetrical styles."
End Sub

Private Sub Command11_Click()
    Form14.Show
    Form15.Hide
End Sub

Private Sub Command12_Click()
Form16.Show
Form15.Hide
End Sub

Private Sub Command13_Click()
Form26.Show
Form15.Hide
End Sub

Private Sub Command2_Click()
    Picture1.Visible = False
    Picture5.Visible = True
    Picture5.Cls
    Picture5.Print "Judith Slaying Holofernes"
    Picture5.Print
    Picture5.Print "by Artemisia Gentileschi"
    Picture5.Print
    Picture5.Print "1617 AD"
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
    Picture6.Print "The Ectasy of Saint Teresa"
    Picture6.Print
    Picture6.Print "by Gianlorenzo Bernini"
    Picture6.Print
    Picture6.Print "1622 AD"
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
    Picture7.Print "David"
    Picture7.Print
    Picture7.Print "by Gianlorenzo Bernini"
    Picture7.Print
    Picture7.Print "1623 AD"
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
    Picture8.Print "The Night Watch"
    Picture8.Print
    Picture8.Print "by Rembrandt van Rijn"
    Picture8.Print
    Picture8.Print "1642 AD"
End Sub

Private Sub Command9_Click()
    Picture4.Visible = True
    Picture8.Visible = False
    Picture8.Cls
End Sub

Private Sub Picture9_Click()
Picture9.Visible = False
End Sub
