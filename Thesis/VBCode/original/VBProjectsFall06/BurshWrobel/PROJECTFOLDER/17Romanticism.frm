VERSION 5.00
Begin VB.Form Form18 
   Caption         =   "Form18"
   ClientHeight    =   8835
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11085
   LinkTopic       =   "Form18"
   Picture         =   "17Romanticism.frx":0000
   ScaleHeight     =   8835
   ScaleWidth      =   11085
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command11 
      Caption         =   "Fav"
      Height          =   375
      Left            =   1320
      TabIndex        =   19
      Top             =   8280
      Width           =   495
   End
   Begin VB.PictureBox Picture7 
      Height          =   1215
      Left            =   3000
      ScaleHeight     =   1155
      ScaleWidth      =   5475
      TabIndex        =   18
      Top             =   2880
      Visible         =   0   'False
      Width           =   5535
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
      Caption         =   "< - Previous"
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
      Height          =   3015
      Left            =   8040
      ScaleHeight     =   2955
      ScaleWidth      =   2835
      TabIndex        =   14
      Top             =   4200
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.PictureBox Picture5 
      Height          =   3015
      Left            =   4080
      ScaleHeight     =   2955
      ScaleWidth      =   3795
      TabIndex        =   13
      Top             =   5160
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.PictureBox Picture4 
      Height          =   3015
      Left            =   120
      ScaleHeight     =   2955
      ScaleWidth      =   3795
      TabIndex        =   12
      Top             =   4200
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.CommandButton Command7 
      Caption         =   "C"
      Height          =   255
      Left            =   10200
      TabIndex        =   11
      Top             =   7320
      Width           =   375
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Info"
      Height          =   255
      Left            =   8400
      TabIndex        =   10
      Top             =   7320
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "C"
      Height          =   255
      Left            =   6720
      TabIndex        =   9
      Top             =   8280
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Info"
      Height          =   255
      Left            =   4680
      TabIndex        =   8
      Top             =   8280
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "C"
      Height          =   255
      Left            =   2520
      TabIndex        =   7
      Top             =   7320
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Info"
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   7320
      Width           =   1695
   End
   Begin VB.PictureBox Picture3 
      Height          =   3015
      Left            =   8040
      Picture         =   "17Romanticism.frx":DC68
      ScaleHeight     =   2955
      ScaleWidth      =   2835
      TabIndex        =   5
      Top             =   4200
      Width           =   2895
   End
   Begin VB.PictureBox Picture2 
      Height          =   3015
      Left            =   4080
      Picture         =   "17Romanticism.frx":11036
      ScaleHeight     =   2955
      ScaleWidth      =   3795
      TabIndex        =   4
      Top             =   5160
      Width           =   3855
   End
   Begin VB.PictureBox Picture1 
      Height          =   3015
      Left            =   120
      Picture         =   "17Romanticism.frx":141AF
      ScaleHeight     =   2955
      ScaleWidth      =   3795
      TabIndex        =   3
      Top             =   4200
      Width           =   3855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Main Menu"
      Height          =   500
      Left            =   240
      TabIndex        =   2
      Top             =   8160
      Width           =   1000
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1800 AD ~ 1850 AD"
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
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   11055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Romanticism"
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
      Left            =   -360
      TabIndex        =   0
      Top             =   120
      Width           =   12975
   End
End
Attribute VB_Name = "Form18"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'WesternArt Project
'Form18
'Bursh,Wrobel
'11-1-06
'This is our Romanticism Form Era, displaying works from Era.
Option Explicit
Private Sub Command1_Click()
    Form1.Show
    Form18.Hide
End Sub

Private Sub Command10_Click()
Form19.Show
Form18.Hide
End Sub

Private Sub Command11_Click()
    Form26.Show
    Form18.Hide
End Sub

Private Sub Command2_Click()
    Picture1.Visible = False
    Picture4.Visible = True
    Picture4.Cls
    Picture4.Print "The Executions of the Third of May,1808"
    Picture4.Print
    Picture4.Print "by Francisco de Goya"
    Picture4.Print
    Picture4.Print "1814 AD"
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
    Picture5.Print "Raft of the Medusa"
    Picture5.Print
    Picture5.Print "by Théodore Gericault"
    Picture5.Print
    Picture5.Print "1819 AD"
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
    Picture6.Print "Liberty Leading the People"
    Picture6.Print
    Picture6.Print "by Eugéne Delecroix"
    Picture6.Print
    Picture6.Print "1830 AD"
End Sub

Private Sub Command7_Click()
    Picture3.Visible = True
    Picture6.Visible = False
    Picture6.Cls
End Sub

Private Sub Command8_Click()
    Picture7.Visible = True
    Picture7.Cls
    Picture7.Print "    With roots coming from the 1700's Romanists were interested in"
    Picture7.Print "the mind as the cause of mysterious, unexplained phenomena.  Artwork"
    Picture7.Print "dealing with psychology and the passing of time are displayed within"
    Picture7.Print "Romantic imagery."
    
End Sub

Private Sub Command9_Click()
Form17.Show
Form18.Hide
End Sub

Private Sub Picture7_Click()
Picture7.Visible = False
End Sub
