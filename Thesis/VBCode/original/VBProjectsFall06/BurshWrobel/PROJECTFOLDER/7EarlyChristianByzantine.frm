VERSION 5.00
Begin VB.Form Form8 
   Caption         =   "Form8"
   ClientHeight    =   8835
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11085
   LinkTopic       =   "Form8"
   Picture         =   "7EarlyChristianByzantine.frx":0000
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
      Height          =   1095
      Left            =   1920
      ScaleHeight     =   1035
      ScaleWidth      =   6795
      TabIndex        =   18
      Top             =   3120
      Visible         =   0   'False
      Width           =   6855
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Next ->"
      Height          =   615
      Left            =   9840
      TabIndex        =   17
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      Caption         =   "<- Previous"
      Height          =   615
      Left            =   0
      TabIndex        =   16
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Introduction to Era"
      Height          =   375
      Left            =   4200
      TabIndex        =   15
      Top             =   2400
      Width           =   2295
   End
   Begin VB.PictureBox Picture6 
      Height          =   2055
      Left            =   7200
      ScaleHeight     =   1995
      ScaleWidth      =   3195
      TabIndex        =   14
      Top             =   6000
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.PictureBox Picture5 
      Height          =   2775
      Left            =   3720
      ScaleHeight     =   2715
      ScaleWidth      =   2835
      TabIndex        =   13
      Top             =   5520
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.PictureBox Picture4 
      Height          =   2775
      Left            =   240
      ScaleHeight     =   2715
      ScaleWidth      =   2835
      TabIndex        =   12
      Top             =   4920
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton Command7 
      Caption         =   "C"
      Height          =   255
      Left            =   9360
      TabIndex        =   11
      Top             =   8160
      Width           =   375
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Info"
      Height          =   255
      Left            =   7560
      TabIndex        =   10
      Top             =   8160
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "C"
      Height          =   255
      Left            =   5760
      TabIndex        =   9
      Top             =   8400
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Info"
      Height          =   255
      Left            =   3960
      TabIndex        =   8
      Top             =   8400
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "C"
      Height          =   255
      Left            =   2280
      TabIndex        =   7
      Top             =   7800
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Info"
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   7800
      Width           =   1695
   End
   Begin VB.PictureBox Picture3 
      Height          =   2055
      Left            =   7200
      Picture         =   "7EarlyChristianByzantine.frx":16110
      ScaleHeight     =   1995
      ScaleWidth      =   3195
      TabIndex        =   5
      Top             =   6000
      Width           =   3255
   End
   Begin VB.PictureBox Picture2 
      Height          =   2775
      Left            =   3720
      Picture         =   "7EarlyChristianByzantine.frx":1A5B4
      ScaleHeight     =   2715
      ScaleWidth      =   2835
      TabIndex        =   4
      Top             =   5520
      Width           =   2895
   End
   Begin VB.PictureBox Picture1 
      Height          =   2775
      Left            =   240
      Picture         =   "7EarlyChristianByzantine.frx":1D21B
      ScaleHeight     =   2715
      ScaleWidth      =   2835
      TabIndex        =   3
      Top             =   4920
      Width           =   2895
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
      Caption         =   "475 AD ~850 AD"
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
      Left            =   -120
      TabIndex        =   1
      Top             =   1560
      Width           =   11055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Early Christian "
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
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   11055
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'WesternArt Project
'Form8
'Bursh,Wrobel
'11-1-06
'This is our Early Christian Form Era, displaying works from Era.
Private Sub Command1_Click()
    Form1.Show
    Form8.Hide
End Sub

Private Sub Command10_Click()
Form9.Show
Form8.Hide
End Sub

Private Sub Command11_Click()
    Form26.Show
    Form8.Hide
End Sub

Private Sub Command2_Click()
    Picture1.Visible = False
    Picture4.Visible = True
    Picture4.Cls
    Picture4.Print "The Hagia Sophia"
    Picture4.Print
    Picture4.Print "537 AD"
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
    Picture5.Print "The San Vitale"
    Picture5.Print
    Picture5.Print "545 AD"
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
    Picture6.Print "The Court of Justinian"
    Picture6.Print
    Picture6.Print "547 AD"
End Sub

Private Sub Command7_Click()
    Picture3.Visible = True
    Picture6.Visible = False
    Picture6.Cls
End Sub

Private Sub Command8_Click()
    Picture7.Visible = True
    Picture7.Cls
    Picture7.Print "    The teaching of Christ and his followers impacted the Western world with promise of eternal"
    Picture7.Print "salvation.  A shift in worshipping style from emperor or multiple Gods to a single divity"
    Picture7.Print "allowed for the construction of Churches and mosaics to express these beliefs and values."
End Sub

Private Sub Command9_Click()
Form7.Show
Form8.Hide
End Sub

Private Sub Picture7_Click()
    Picture7.Visible = False
    
End Sub
