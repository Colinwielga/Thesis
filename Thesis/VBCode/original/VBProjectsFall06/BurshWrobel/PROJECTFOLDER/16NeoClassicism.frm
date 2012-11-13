VERSION 5.00
Begin VB.Form Form17 
   Caption         =   "Form17"
   ClientHeight    =   8835
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11085
   LinkTopic       =   "Form17"
   Picture         =   "16NeoClassicism.frx":0000
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
      Left            =   2760
      ScaleHeight     =   1035
      ScaleWidth      =   5355
      TabIndex        =   18
      Top             =   3000
      Visible         =   0   'False
      Width           =   5415
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
      Top             =   2280
      Width           =   2295
   End
   Begin VB.PictureBox Picture6 
      Height          =   3735
      Left            =   8040
      ScaleHeight     =   3675
      ScaleWidth      =   2475
      TabIndex        =   14
      Top             =   4560
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.PictureBox Picture5 
      Height          =   3735
      Left            =   4920
      ScaleHeight     =   3675
      ScaleWidth      =   2475
      TabIndex        =   13
      Top             =   4560
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.PictureBox Picture4 
      Height          =   2895
      Left            =   480
      ScaleHeight     =   2835
      ScaleWidth      =   3795
      TabIndex        =   12
      Top             =   4680
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.CommandButton Command7 
      Caption         =   "C"
      Height          =   255
      Left            =   9960
      TabIndex        =   11
      Top             =   8400
      Width           =   375
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Info"
      Height          =   255
      Left            =   8160
      TabIndex        =   10
      Top             =   8400
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "C"
      Height          =   255
      Left            =   6840
      TabIndex        =   9
      Top             =   8400
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Info"
      Height          =   255
      Left            =   5040
      TabIndex        =   8
      Top             =   8400
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "C"
      Height          =   255
      Left            =   2880
      TabIndex        =   7
      Top             =   7680
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Info"
      Height          =   255
      Left            =   1080
      TabIndex        =   6
      Top             =   7680
      Width           =   1695
   End
   Begin VB.PictureBox Picture3 
      Height          =   3735
      Left            =   8040
      Picture         =   "16NeoClassicism.frx":16DC3
      ScaleHeight     =   3675
      ScaleWidth      =   2475
      TabIndex        =   5
      Top             =   4560
      Width           =   2535
   End
   Begin VB.PictureBox Picture2 
      Height          =   3735
      Left            =   4920
      Picture         =   "16NeoClassicism.frx":1A7B0
      ScaleHeight     =   3675
      ScaleWidth      =   2475
      TabIndex        =   4
      Top             =   4560
      Width           =   2535
   End
   Begin VB.PictureBox Picture1 
      Height          =   2895
      Left            =   480
      Picture         =   "16NeoClassicism.frx":1C3DA
      ScaleHeight     =   2835
      ScaleWidth      =   3795
      TabIndex        =   3
      Top             =   4680
      Width           =   3855
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
      Caption         =   "1750 AD ~ 1800 AD"
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
      Caption         =   "NeoClassicism"
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
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   11055
   End
End
Attribute VB_Name = "Form17"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'WesternArt Project
'Form17
'Bursh,Wrobel
'11-1-06
'This is our NeoClassicism Form Era, displaying works from Era.
Option Explicit
Private Sub Command1_Click()
    Form1.Show
    Form17.Hide
End Sub

Private Sub Command10_Click()
Form18.Show
Form17.Hide
End Sub

Private Sub Command11_Click()
Form26.Show
Form17.Hide
End Sub

Private Sub Command2_Click()
    Picture1.Visible = False
    Picture4.Visible = True
    Picture4.Cls
    Picture4.Print "Oath of the Horatii"
    Picture4.Print
    Picture4.Print "by Jacques-Louis David"
    Picture4.Print
    Picture4.Print "1785 AD"
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
    Picture5.Print "Death of Morat"
    Picture5.Print
    Picture5.Print "by Jacques-Louis David"
    Picture5.Print
    Picture5.Print "1793 AD"
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
    Picture6.Print "Napoleon Enthroned"
    Picture6.Print
    Picture6.Print "by Jean-Auguste Dominique Ingres"
    Picture6.Print
    Picture6.Print "1800 AD"
End Sub

Private Sub Command7_Click()
    Picture3.Visible = True
    Picture6.Visible = False
    Picture6.Cls
End Sub

Private Sub Command8_Click()
    Picture7.Visible = True
    Picture7.Cls
    Picture7.Print "    With several styles competing for dominance in France the Neoclassic"
    Picture7.Print "style -truestyle- was a reaction against Rococo levity, and can be"
    Picture7.Print "associated with military movements of The French Revolution with a"
    Picture7.Print "Classical flavor."

End Sub

Private Sub Command9_Click()
    Form16.Show
    Form17.Hide
End Sub

Private Sub Picture7_Click()
    Picture7.Visible = False
End Sub
