VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "Form6"
   ClientHeight    =   8835
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11085
   LinkTopic       =   "Form6"
   Picture         =   "5HellenisticGreece.frx":0000
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
      Left            =   2040
      ScaleHeight     =   1035
      ScaleWidth      =   7155
      TabIndex        =   18
      Top             =   3240
      Visible         =   0   'False
      Width           =   7215
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
      Left            =   0
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
      Height          =   2775
      Left            =   8160
      ScaleHeight     =   2715
      ScaleWidth      =   2115
      TabIndex        =   14
      Top             =   4680
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.PictureBox Picture5 
      Height          =   2775
      Left            =   3840
      ScaleHeight     =   2715
      ScaleWidth      =   3315
      TabIndex        =   13
      Top             =   5160
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.PictureBox Picture4 
      Height          =   2775
      Left            =   480
      ScaleHeight     =   2715
      ScaleWidth      =   2115
      TabIndex        =   12
      Top             =   4560
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton Command7 
      Caption         =   "C"
      Height          =   255
      Left            =   9840
      TabIndex        =   11
      Top             =   7560
      Width           =   375
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Info"
      Height          =   255
      Left            =   8040
      TabIndex        =   10
      Top             =   7560
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "C"
      Height          =   255
      Left            =   6240
      TabIndex        =   9
      Top             =   8160
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Info"
      Height          =   255
      Left            =   4320
      TabIndex        =   8
      Top             =   8160
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "C"
      Height          =   255
      Left            =   2280
      TabIndex        =   7
      Top             =   7440
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Info"
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   7440
      Width           =   1695
   End
   Begin VB.PictureBox Picture3 
      Height          =   2775
      Left            =   8160
      Picture         =   "5HellenisticGreece.frx":13C0E
      ScaleHeight     =   2715
      ScaleWidth      =   2115
      TabIndex        =   5
      Top             =   4680
      Width           =   2175
   End
   Begin VB.PictureBox Picture2 
      Height          =   2775
      Left            =   3840
      Picture         =   "5HellenisticGreece.frx":15829
      ScaleHeight     =   2715
      ScaleWidth      =   3315
      TabIndex        =   4
      Top             =   5160
      Width           =   3375
   End
   Begin VB.PictureBox Picture1 
      Height          =   2775
      Left            =   480
      Picture         =   "5HellenisticGreece.frx":187DB
      ScaleHeight     =   2715
      ScaleWidth      =   2115
      TabIndex        =   3
      Top             =   4560
      Width           =   2175
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
      Caption         =   "350 BCE ~ 0"
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
      Left            =   0
      TabIndex        =   1
      Top             =   1920
      Width           =   11055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Hellenistic Greece"
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
      Top             =   360
      Width           =   11055
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'WesternArt Project
'Form6
'Bursh,Wrobel
'11-1-06
'This is our Hellenistic Form Era, displaying works from Era.
Private Sub Command1_Click()
    Form1.Show
    Form6.Hide
End Sub

Private Sub Command10_Click()
    Form7.Show
    Form6.Hide
End Sub

Private Sub Command11_Click()
Form26.Show
Form6.Hide
End Sub

Private Sub Command2_Click()
Picture1.Visible = False
    Picture4.Visible = True
    Picture4.Cls
    Picture4.Print "Winged Nike"
    Picture4.Print
    Picture4.Print "190 BCE"
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
    Picture5.Print "Athena Battling with Alkyoneus"
    Picture5.Print
    Picture5.Print "180 BCE"
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
    Picture6.Print "Laocoòn and His Two Sons"
    Picture6.Print
    Picture6.Print "30 BCE"
End Sub

Private Sub Command7_Click()
    Picture3.Visible = True
    Picture6.Visible = False
    Picture6.Cls
End Sub

Private Sub Command8_Click()
    Picture7.Visible = True
    Picture7.Cls
    Picture7.Print "    Hellenistic refers to the spread of Greek culture beyond Greece.  There is an increase in "
    Picture7.Print "variety of subject matter, in particular the nature of childhood.  An exceptional amount of "
    Picture7.Print "emotion is expressed through realistic character actions."
End Sub

Private Sub Command9_Click()
    Form5.Show
    Form6.Hide
End Sub

Private Sub Picture7_Click()
    Picture7.Visible = False
End Sub
