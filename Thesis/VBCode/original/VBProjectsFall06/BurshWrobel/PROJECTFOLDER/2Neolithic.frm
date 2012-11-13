VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   8835
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11085
   LinkTopic       =   "Form3"
   Picture         =   "2Neolithic.frx":0000
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
      Height          =   975
      Left            =   2160
      ScaleHeight     =   915
      ScaleWidth      =   6675
      TabIndex        =   18
      Top             =   3000
      Visible         =   0   'False
      Width           =   6735
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
      Left            =   4320
      TabIndex        =   15
      Top             =   2520
      Width           =   2295
   End
   Begin VB.PictureBox Picture6 
      Height          =   2775
      Left            =   7680
      ScaleHeight     =   2715
      ScaleWidth      =   2235
      TabIndex        =   14
      Top             =   4200
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.PictureBox Picture5 
      Height          =   2775
      Left            =   4320
      ScaleHeight     =   2715
      ScaleWidth      =   2235
      TabIndex        =   13
      Top             =   4200
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.PictureBox Picture4 
      Height          =   2775
      Left            =   720
      ScaleHeight     =   2715
      ScaleWidth      =   2235
      TabIndex        =   12
      Top             =   4200
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton Command7 
      Caption         =   "C"
      Height          =   255
      Left            =   9480
      TabIndex        =   11
      Top             =   7080
      Width           =   495
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Info"
      Height          =   255
      Left            =   7680
      TabIndex        =   10
      Top             =   7080
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "C"
      Height          =   255
      Left            =   6120
      TabIndex        =   9
      Top             =   7080
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Info"
      Height          =   255
      Left            =   4320
      TabIndex        =   8
      Top             =   7080
      Width           =   1655
   End
   Begin VB.CommandButton Command3 
      Caption         =   "C"
      Height          =   255
      Left            =   2640
      TabIndex        =   7
      Top             =   7080
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Info"
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   7080
      Width           =   1655
   End
   Begin VB.PictureBox Picture3 
      Height          =   2775
      Left            =   7680
      Picture         =   "2Neolithic.frx":19FD0
      ScaleHeight     =   2715
      ScaleWidth      =   2220
      TabIndex        =   5
      Top             =   4200
      Width           =   2280
   End
   Begin VB.PictureBox Picture2 
      Height          =   2775
      Left            =   4320
      Picture         =   "2Neolithic.frx":1C1AE
      ScaleHeight     =   2715
      ScaleWidth      =   2235
      TabIndex        =   4
      Top             =   4200
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      Height          =   2775
      Index           =   0
      Left            =   720
      Picture         =   "2Neolithic.frx":1F527
      ScaleHeight     =   2715
      ScaleWidth      =   2235
      TabIndex        =   3
      Top             =   4200
      Width           =   2295
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
      Caption         =   "6,000 BCE ~ 2,500 BCE"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   23.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   10935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Neolithic Era"
      BeginProperty Font 
         Name            =   "High Tower Text"
         Size            =   69
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   10935
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'WesternArt Project
'Form3
'Bursh,Wrobel
'11-1-06
'This is our Neolithic Form Era, displaying works from Era.
Option Explicit

Private Sub Command1_Click()
    Form1.Show
    Form3.Hide
End Sub


Private Sub Command10_Click()
    Form4.Show
    Form3.Hide
End Sub

Private Sub Command11_Click()
Form26.Show
Form3.Hide
End Sub

Private Sub Command2_Click()
    Picture1(0).Visible = False
    Picture4.Visible = True
    Picture4.Cls
    Picture4.Print "Menhirs"
    Picture4.Print
    Picture4.Print "4,000 BCE"
End Sub

Private Sub Command3_Click()
    Picture1(0).Visible = True
    Picture4.Visible = False
    Picture4.Cls
End Sub

Private Sub Command4_Click()
    Picture2.Visible = False
    Picture5.Visible = True
    Picture5.Cls
    Picture5.Print "Dolmen"
    Picture5.Print
    Picture5.Print "4,000 BCE"
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
    Picture6.Print "Cromlechs"
    Picture6.Print
    Picture6.Print "2,800 BCE"
End Sub

Private Sub Command7_Click()
    Picture3.Visible = True
    Picture6.Visible = False
    Picture6.Cls
End Sub

Private Sub Command8_Click()
    Picture7.Visible = True
    Picture7.Cls
    Picture7.Print "    The transition of hunting and gathering to a farming agriculture lead to the development of"
    Picture7.Print "monumental stone architecture; in megaliths.  These megaliths having religious associations and"
    Picture7.Print "death rituals were made with intention of out lasting any mortals lifespan."
End Sub

Private Sub Command9_Click()
    Form2.Show
    Form3.Hide
End Sub

Private Sub Picture7_Click()
    Picture7.Visible = False
End Sub
