VERSION 5.00
Begin VB.Form Form7 
   Caption         =   "Form7"
   ClientHeight    =   8835
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11085
   LinkTopic       =   "Form7"
   Picture         =   "6AncientRome.frx":0000
   ScaleHeight     =   8835
   ScaleWidth      =   11085
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton Command11 
      Caption         =   "Fav"
      Height          =   375
      Left            =   1440
      TabIndex        =   19
      Top             =   8160
      Width           =   495
   End
   Begin VB.PictureBox Picture7 
      Height          =   1215
      Left            =   2160
      ScaleHeight     =   1155
      ScaleWidth      =   7035
      TabIndex        =   18
      Top             =   3360
      Visible         =   0   'False
      Width           =   7095
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Next ->"
      Height          =   615
      Left            =   9960
      TabIndex        =   17
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      Caption         =   "<- Previous"
      Height          =   615
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Introduction to Era"
      Height          =   375
      Left            =   4440
      TabIndex        =   15
      Top             =   2640
      Width           =   2295
   End
   Begin VB.PictureBox Picture6 
      Height          =   2295
      Left            =   7560
      ScaleHeight     =   2235
      ScaleWidth      =   3075
      TabIndex        =   14
      Top             =   5640
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.PictureBox Picture5 
      Height          =   2295
      Left            =   3960
      ScaleHeight     =   2235
      ScaleWidth      =   3075
      TabIndex        =   13
      Top             =   5520
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.PictureBox Picture4 
      Height          =   2295
      Left            =   360
      ScaleHeight     =   2235
      ScaleWidth      =   3075
      TabIndex        =   12
      Top             =   5040
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton Command7 
      Caption         =   "C"
      Height          =   255
      Left            =   9720
      TabIndex        =   11
      Top             =   8040
      Width           =   375
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Info"
      Height          =   255
      Left            =   7920
      TabIndex        =   10
      Top             =   8040
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "C"
      Height          =   255
      Left            =   6000
      TabIndex        =   9
      Top             =   7920
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Info"
      Height          =   255
      Left            =   4200
      TabIndex        =   8
      Top             =   7920
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "C"
      Height          =   255
      Left            =   2520
      TabIndex        =   7
      Top             =   7440
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Info"
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   7440
      Width           =   1695
   End
   Begin VB.PictureBox Picture3 
      Height          =   2295
      Left            =   7560
      Picture         =   "6AncientRome.frx":1B446
      ScaleHeight     =   2235
      ScaleWidth      =   3105
      TabIndex        =   5
      Top             =   5640
      Width           =   3165
   End
   Begin VB.PictureBox Picture2 
      Height          =   2295
      Left            =   3960
      Picture         =   "6AncientRome.frx":1D909
      ScaleHeight     =   2235
      ScaleWidth      =   3075
      TabIndex        =   4
      Top             =   5520
      Width           =   3135
   End
   Begin VB.PictureBox Picture1 
      Height          =   2295
      Left            =   360
      Picture         =   "6AncientRome.frx":1F7F9
      ScaleHeight     =   2235
      ScaleWidth      =   3075
      TabIndex        =   3
      Top             =   5040
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Main Menu"
      Height          =   500
      Left            =   360
      TabIndex        =   2
      Top             =   8040
      Width           =   1000
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0 ~ 475 AD"
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
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   10935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ancient Rome"
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
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   10935
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'WesternArt Project
'Form7
'Bursh,Wrobel
'11-1-06
'This is our Ancient Rome Form Era, displaying works from Era.
Option Explicit
Private Sub Command1_Click()
    Form1.Show
    Form7.Hide
End Sub

Private Sub Command10_Click()
    Form8.Show
    Form7.Hide
End Sub

Private Sub Command11_Click()
Form26.Show
Form7.Hide
End Sub

Private Sub Command2_Click()
Picture1.Visible = False
    Picture4.Visible = True
    Picture4.Cls
    Picture4.Print "The Colosseum"
    Picture4.Print
    Picture4.Print "75 AD"
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
    Picture5.Print "The Pantheon"
    Picture5.Print
    Picture5.Print "120 AD"
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
    Picture6.Print "The Arch of Constantine"
    Picture6.Print
    Picture6.Print "313 AD"
End Sub

Private Sub Command7_Click()
    Picture3.Visible = True
    Picture6.Visible = False
    Picture6.Cls
End Sub

Private Sub Command8_Click()
    Picture7.Visible = True
    Picture7.Cls
    Picture7.Print "    Romes designation as - caput mundi - 'capital of the world' claimed its place at the center"
    Picture7.Print "of world power.  Influenced greatly by Greek works, Roman art was typically narrative or based on"
    Picture7.Print "actual historical events - connecting the present with the past.  With the extensive use of concrete"
    Picture7.Print "Roman architecture showed great praise for its citizens, its state, and its emperor."
End Sub

Private Sub Command9_Click()
Form6.Show
Form7.Hide
End Sub

Private Sub Picture7_Click()
    Picture7.Visible = False
End Sub
