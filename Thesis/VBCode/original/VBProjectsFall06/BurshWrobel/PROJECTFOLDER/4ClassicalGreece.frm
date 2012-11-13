VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Form5"
   ClientHeight    =   8835
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11085
   LinkTopic       =   "Form5"
   Picture         =   "4ClassicalGreece.frx":0000
   ScaleHeight     =   8835
   ScaleWidth      =   11085
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command11 
      Caption         =   "Fav"
      Height          =   375
      Left            =   1440
      TabIndex        =   19
      Top             =   8280
      Width           =   495
   End
   Begin VB.PictureBox Picture7 
      Height          =   1215
      Left            =   2040
      ScaleHeight     =   1155
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
      Top             =   2760
      Width           =   2295
   End
   Begin VB.PictureBox Picture6 
      Height          =   2535
      Left            =   7200
      ScaleHeight     =   2475
      ScaleWidth      =   3315
      TabIndex        =   14
      Top             =   4920
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.PictureBox Picture5 
      Height          =   2655
      Left            =   4200
      ScaleHeight     =   2595
      ScaleWidth      =   2115
      TabIndex        =   13
      Top             =   4800
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton Command7 
      Caption         =   "C"
      Height          =   255
      Left            =   9240
      TabIndex        =   12
      Top             =   7560
      Width           =   375
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Info"
      Height          =   255
      Left            =   7560
      TabIndex        =   11
      Top             =   7560
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "C"
      Height          =   255
      Left            =   6000
      TabIndex        =   10
      Top             =   7560
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Info"
      Height          =   255
      Left            =   4200
      TabIndex        =   9
      Top             =   7560
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "C"
      Height          =   255
      Left            =   3000
      TabIndex        =   8
      Top             =   7560
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Info"
      Height          =   255
      Left            =   1200
      TabIndex        =   7
      Top             =   7560
      Width           =   1695
   End
   Begin VB.PictureBox Picture4 
      Height          =   2655
      Left            =   1200
      ScaleHeight     =   2595
      ScaleWidth      =   2115
      TabIndex        =   6
      Top             =   4800
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.PictureBox Picture3 
      Height          =   2535
      Left            =   7200
      Picture         =   "4ClassicalGreece.frx":2C2AB
      ScaleHeight     =   2475
      ScaleWidth      =   3315
      TabIndex        =   5
      Top             =   4920
      Width           =   3375
   End
   Begin VB.PictureBox Picture2 
      Height          =   2655
      Left            =   4200
      Picture         =   "4ClassicalGreece.frx":2E1AD
      ScaleHeight     =   2595
      ScaleWidth      =   2115
      TabIndex        =   4
      Top             =   4800
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      Height          =   2655
      Left            =   1200
      Picture         =   "4ClassicalGreece.frx":307B3
      ScaleHeight     =   2595
      ScaleWidth      =   2115
      TabIndex        =   3
      Top             =   4800
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Main Menu"
      Height          =   500
      Left            =   360
      TabIndex        =   2
      Top             =   8160
      Width           =   1000
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "450 BCE ~ 350 BCE"
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
      Top             =   1920
      Width           =   11055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Classical Greece"
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
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   10935
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'WesternArt Project
'Form5
'Bursh,Wrobel
'11-1-06
'This is our Classical Greece Form Era, displaying works from Era.
Option Explicit
Private Sub Command1_Click()
    Form1.Show
    Form5.Hide
End Sub

Private Sub Command10_Click()
    Form6.Show
    Form5.Hide
End Sub

Private Sub Command11_Click()
Form26.Show
Form5.Hide
End Sub

Private Sub Command2_Click()
    Picture1.Visible = False
    Picture4.Visible = True
    Picture4.Cls
    Picture4.Print "Warrior From Riace"
    Picture4.Print
    Picture4.Print "450 BCE"
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
    Picture5.Print "Polykleitos of Argos"
    Picture5.Print
    Picture5.Print "440 BCE"
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
    Picture6.Print "The Parthenon"
    Picture6.Print
    Picture6.Print "440 BCE"
End Sub

Private Sub Command7_Click()
    Picture3.Visible = True
    Picture6.Visible = False
    Picture6.Cls
End Sub

Private Sub Command8_Click()
    Picture7.Visible = True
    Picture7.Cls
    Picture7.Print "    Damed as the 'Golden Age' of Greek art, the Classical period refferd to significant Greek cultural"
    Picture7.Print "and intellectual accomplishments of the 5th Century BCE.  The common phrase 'Man is the Measure "
    Picture7.Print "of all Things' put emphasis on individual strengths and philosophies."
End Sub

Private Sub Command9_Click()
    Form4.Show
    Form5.Hide
End Sub

Private Sub Picture7_Click()
    Picture7.Visible = False
End Sub
