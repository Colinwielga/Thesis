VERSION 5.00
Begin VB.Form Form19 
   Caption         =   "Form19"
   ClientHeight    =   8835
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11085
   LinkTopic       =   "Form19"
   Picture         =   "18Realism.frx":0000
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
      Left            =   5040
      ScaleHeight     =   1155
      ScaleWidth      =   4755
      TabIndex        =   18
      Top             =   3240
      Visible         =   0   'False
      Width           =   4815
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
      Top             =   2400
      Width           =   2295
   End
   Begin VB.PictureBox Picture6 
      Height          =   2775
      Left            =   7200
      ScaleHeight     =   2715
      ScaleWidth      =   3555
      TabIndex        =   14
      Top             =   5640
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.PictureBox Picture5 
      Height          =   2775
      Left            =   3120
      ScaleHeight     =   2715
      ScaleWidth      =   3675
      TabIndex        =   13
      Top             =   5640
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.PictureBox Picture4 
      Height          =   2295
      Left            =   120
      ScaleHeight     =   2235
      ScaleWidth      =   4035
      TabIndex        =   12
      Top             =   3120
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.CommandButton Command7 
      Caption         =   "C"
      Height          =   255
      Left            =   9600
      TabIndex        =   11
      Top             =   8520
      Width           =   375
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Info"
      Height          =   255
      Left            =   7800
      TabIndex        =   10
      Top             =   8520
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "C"
      Height          =   255
      Left            =   5640
      TabIndex        =   9
      Top             =   8520
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Info"
      Height          =   255
      Left            =   3840
      TabIndex        =   8
      Top             =   8520
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "C"
      Height          =   255
      Left            =   2520
      TabIndex        =   7
      Top             =   5520
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Info"
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   5520
      Width           =   1695
   End
   Begin VB.PictureBox Picture3 
      Height          =   2775
      Left            =   7200
      Picture         =   "18Realism.frx":FC1E
      ScaleHeight     =   2715
      ScaleWidth      =   3555
      TabIndex        =   5
      Top             =   5640
      Width           =   3615
   End
   Begin VB.PictureBox Picture2 
      Height          =   2775
      Left            =   3120
      Picture         =   "18Realism.frx":12CE9
      ScaleHeight     =   2715
      ScaleWidth      =   3675
      TabIndex        =   4
      Top             =   5640
      Width           =   3735
   End
   Begin VB.PictureBox Picture1 
      Height          =   2295
      Left            =   120
      Picture         =   "18Realism.frx":15022
      ScaleHeight     =   2235
      ScaleWidth      =   4035
      TabIndex        =   3
      Top             =   3120
      Width           =   4095
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
      Caption         =   "1850 AD ~ 1875 AD"
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
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   10935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Realism"
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
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   10935
   End
End
Attribute VB_Name = "Form19"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'WesternArt Project
'Form19
'Bursh,Wrobel
'11-1-06
'This is our Realism Form Era, displaying works from Era.
Option Explicit
Private Sub Command1_Click()
    Form1.Show
    Form19.Hide
End Sub

Private Sub Command10_Click()
Form20.Show
Form19.Hide
End Sub

Private Sub Command11_Click()
    Form26.Show
    Form19.Hide
End Sub

Private Sub Command2_Click()
    Picture1.Visible = False
    Picture4.Visible = True
    Picture4.Cls
    Picture4.Print "The Interior of My Studio: A Real Allegory Summoning Up"
    Picture4.Print "Seven Years of My Life as an Artist from 1848 to 1855 "
    Picture4.Print
    Picture4.Print "by Gustave Courbet"
    Picture4.Print
    Picture4.Print "1855 AD"
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
    Picture5.Print "Third-Class Carriage"
    Picture5.Print
    Picture5.Print "by Honoré Daumier"
    Picture5.Print
    Picture5.Print "1862 AD"
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
    Picture6.Print "Le Déjeuner sur l'Herbe"
    Picture6.Print
    Picture6.Print "by Édouard Manet"
    Picture6.Print
    Picture6.Print "1863 AD"
End Sub

Private Sub Command7_Click()
    Picture3.Visible = True
    Picture6.Visible = False
    Picture6.Cls
End Sub

Private Sub Command8_Click()
    Picture7.Visible = True
    Picture7.Cls
    Picture7.Print "    With the Age of Revolution, realistic art displayed the social"
    Picture7.Print "and economic situations of the world through imagery.  Works that"
    Picture7.Print "displayed this direct observation of society and social awareness"
    Picture7.Print "were subject to a great amount of criticism."
End Sub

Private Sub Command9_Click()
Form18.Show
Form19.Hide
End Sub

Private Sub Picture7_Click()
Picture7.Visible = False
End Sub
