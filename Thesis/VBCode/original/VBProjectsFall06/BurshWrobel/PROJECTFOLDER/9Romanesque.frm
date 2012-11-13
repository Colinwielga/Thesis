VERSION 5.00
Begin VB.Form Form10 
   BackColor       =   &H00000000&
   Caption         =   "Form10"
   ClientHeight    =   8835
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11085
   LinkTopic       =   "Form10"
   Picture         =   "9Romanesque.frx":0000
   ScaleHeight     =   8835
   ScaleWidth      =   11085
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command11 
      Caption         =   "Fav"
      Height          =   375
      Left            =   1200
      TabIndex        =   19
      Top             =   8160
      Width           =   495
   End
   Begin VB.PictureBox Picture7 
      Height          =   1215
      Left            =   2160
      ScaleHeight     =   1155
      ScaleWidth      =   6555
      TabIndex        =   18
      Top             =   3000
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
      Left            =   4320
      TabIndex        =   15
      Top             =   2280
      Width           =   2295
   End
   Begin VB.CommandButton Command7 
      Caption         =   "C"
      Height          =   255
      Left            =   9480
      TabIndex        =   14
      Top             =   7800
      Width           =   375
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Info"
      Height          =   255
      Left            =   7680
      TabIndex        =   13
      Top             =   7800
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "C"
      Height          =   255
      Left            =   5880
      TabIndex        =   12
      Top             =   7800
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Info"
      Height          =   255
      Left            =   4080
      TabIndex        =   11
      Top             =   7800
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "C"
      Height          =   255
      Left            =   2520
      TabIndex        =   10
      Top             =   7440
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Info"
      Height          =   255
      Left            =   720
      TabIndex        =   9
      Top             =   7440
      Width           =   1695
   End
   Begin VB.PictureBox Picture6 
      Height          =   2055
      Left            =   7200
      ScaleHeight     =   1995
      ScaleWidth      =   3315
      TabIndex        =   8
      Top             =   5640
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.PictureBox Picture5 
      Height          =   2055
      Left            =   3480
      ScaleHeight     =   1995
      ScaleWidth      =   3315
      TabIndex        =   7
      Top             =   5640
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.PictureBox Picture4 
      Height          =   2415
      Left            =   720
      ScaleHeight     =   2355
      ScaleWidth      =   2235
      TabIndex        =   6
      Top             =   4920
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.PictureBox Picture3 
      Height          =   2055
      Left            =   7200
      Picture         =   "9Romanesque.frx":F48E
      ScaleHeight     =   1995
      ScaleWidth      =   3315
      TabIndex        =   5
      Top             =   5640
      Width           =   3375
   End
   Begin VB.PictureBox Picture2 
      Height          =   2055
      Left            =   3480
      Picture         =   "9Romanesque.frx":12E23
      ScaleHeight     =   1995
      ScaleWidth      =   3315
      TabIndex        =   4
      Top             =   5640
      Width           =   3375
   End
   Begin VB.PictureBox Picture1 
      Height          =   2415
      Left            =   720
      Picture         =   "9Romanesque.frx":16ACB
      ScaleHeight     =   2355
      ScaleWidth      =   2235
      TabIndex        =   3
      Top             =   4920
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Main Menu"
      Height          =   500
      Left            =   120
      TabIndex        =   2
      Top             =   8040
      Width           =   1000
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1000 AD ~ 1150 AD"
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
      Left            =   0
      TabIndex        =   1
      Top             =   1560
      Width           =   11055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Romanesque"
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
      Height          =   1935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11055
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'WesternArt Project
'Form10
'Bursh,Wrobel
'11-1-06
'This is our Romanesque Form Era, displaying works from Era.
Option Explicit
Private Sub Command1_Click()
    Form1.Show
    Form10.Hide
End Sub

Private Sub Command10_Click()
    Form11.Show
    Form10.Hide
End Sub

Private Sub Command11_Click()
Form26.Show
Form10.Hide
End Sub

Private Sub Command2_Click()
    Picture1.Visible = False
    Picture4.Visible = True
    Picture4.Cls
    Picture4.Print "Sainte-Foy at Conques"
    Picture4.Print
    Picture4.Print "1075 AD"
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
    Picture5.Print "The Bayeux Tapestry"
    Picture5.Print
    Picture5.Print "1075 AD"
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
    Picture6.Print "The Last Judgement"
    Picture6.Print
    Picture6.Print "845 AD"
End Sub

Private Sub Command7_Click()
    Picture3.Visible = True
    Picture6.Visible = False
    Picture6.Cls
End Sub

Private Sub Command8_Click()
    Picture7.Visible = True
    Picture7.Cls
    Picture7.Print "    Romanesque meaning 'Roman-like' and being associated with architecture with features"
    Picture7.Print "such as round arches, stone vaults, thick walls and exterior reliefs, which reflected the"
    Picture7.Print "stability and prosperity of the Christian Church and all the mural artwork found within."
End Sub

Private Sub Command9_Click()
    Form9.Show
    Form10.Hide
End Sub

Private Sub Picture7_Click()
Picture7.Visible = False
End Sub
