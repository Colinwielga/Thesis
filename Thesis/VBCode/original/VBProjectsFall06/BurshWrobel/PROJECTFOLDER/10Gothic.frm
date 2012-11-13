VERSION 5.00
Begin VB.Form Form11 
   BackColor       =   &H00000000&
   Caption         =   "Form11"
   ClientHeight    =   8835
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11085
   LinkTopic       =   "Form11"
   Picture         =   "10Gothic.frx":0000
   ScaleHeight     =   8835
   ScaleWidth      =   11085
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton Command11 
      Caption         =   "Fav"
      Height          =   375
      Left            =   1200
      TabIndex        =   19
      Top             =   8280
      Width           =   495
   End
   Begin VB.PictureBox Picture7 
      Height          =   1455
      Left            =   2280
      ScaleHeight     =   1395
      ScaleWidth      =   6555
      TabIndex        =   18
      Top             =   2880
      Visible         =   0   'False
      Width           =   6615
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Next ->"
      Height          =   615
      Left            =   9960
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
      Left            =   4560
      TabIndex        =   15
      Top             =   2280
      Width           =   2295
   End
   Begin VB.PictureBox Picture6 
      Height          =   3375
      Left            =   7800
      ScaleHeight     =   3315
      ScaleWidth      =   2475
      TabIndex        =   14
      Top             =   4800
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.PictureBox Picture5 
      Height          =   3375
      Left            =   4560
      ScaleHeight     =   3315
      ScaleWidth      =   2595
      TabIndex        =   13
      Top             =   4800
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.PictureBox Picture4 
      Height          =   2895
      Left            =   120
      ScaleHeight     =   2835
      ScaleWidth      =   3795
      TabIndex        =   12
      Top             =   4560
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.CommandButton Command7 
      Caption         =   "C"
      Height          =   255
      Left            =   9840
      TabIndex        =   11
      Top             =   8280
      Width           =   375
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Info"
      Height          =   255
      Left            =   8040
      TabIndex        =   10
      Top             =   8280
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "C"
      Height          =   255
      Left            =   6600
      TabIndex        =   9
      Top             =   8280
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Info"
      Height          =   255
      Left            =   4800
      TabIndex        =   8
      Top             =   8280
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "C"
      Height          =   255
      Left            =   2520
      TabIndex        =   7
      Top             =   7560
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Info"
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   7560
      Width           =   1695
   End
   Begin VB.PictureBox Picture3 
      Height          =   3375
      Left            =   7800
      Picture         =   "10Gothic.frx":13892
      ScaleHeight     =   3315
      ScaleWidth      =   2475
      TabIndex        =   5
      Top             =   4800
      Width           =   2535
   End
   Begin VB.PictureBox Picture2 
      Height          =   3375
      Left            =   4560
      Picture         =   "10Gothic.frx":179B9
      ScaleHeight     =   3315
      ScaleWidth      =   2595
      TabIndex        =   4
      Top             =   4800
      Width           =   2655
   End
   Begin VB.PictureBox Picture1 
      Height          =   2895
      Left            =   120
      Picture         =   "10Gothic.frx":1B618
      ScaleHeight     =   2835
      ScaleWidth      =   3795
      TabIndex        =   3
      Top             =   4560
      Width           =   3855
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
      Caption         =   "1150 AD ~ 1300 AD"
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
      Caption         =   "Gothic"
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
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   11055
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'WesternArt Project
'Form11
'Bursh,Wrobel
'11-1-06
'This is our Gothic Form Era, displaying works from Era.
Option Explicit
Private Sub Command1_Click()
    Form1.Show
    Form11.Hide
End Sub

Private Sub Command10_Click()
Form12.Show
Form11.Hide
End Sub

Private Sub Command11_Click()
Form26.Show
Form11.Hide
End Sub

Private Sub Command2_Click()
    Picture1.Visible = False
    Picture4.Visible = True
    Picture4.Cls
    Picture4.Print "The Chartes Cathedral"
    Picture4.Print
    Picture4.Print "1150 AD"
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
    Picture5.Print "The Reims Cathedral"
    Picture5.Print
    Picture5.Print "1211 AD"
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
    Picture6.Print "The Salisbury Cathedral"
    Picture6.Print
    Picture6.Print "1220 AD"
End Sub

Private Sub Command7_Click()
    Picture3.Visible = True
    Picture6.Visible = False
    Picture6.Cls
End Sub

Private Sub Command8_Click()
    Picture7.Visible = True
    Picture7.Cls
    Picture7.Print "    Goths were the Germanic tribes who sacked Rome and destroyed the Classical Style,"
    Picture7.Print "and their architecture is meant to represent their lifestyle.  Gothic cathedrals are"
    Picture7.Print "among the greatest stone monuments.  Constructed of rib vaults, flying buttresses, pointed"
    Picture7.Print "arches and covered with stained glass windows - Gothic cathedrals express the relationship"
    Picture7.Print "between light and God in a distinctive way."
End Sub

Private Sub Command9_Click()
 Form10.Show
 Form11.Hide
End Sub


Private Sub Picture7_Click()
Picture7.Visible = False
End Sub
