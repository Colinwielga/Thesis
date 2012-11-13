VERSION 5.00
Begin VB.Form Form12 
   Caption         =   "Form12"
   ClientHeight    =   8835
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11085
   LinkTopic       =   "Form12"
   Picture         =   "11EarlyRenaissance.frx":0000
   ScaleHeight     =   8835
   ScaleWidth      =   11085
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command13 
      Caption         =   "Fav"
      Height          =   375
      Left            =   1320
      TabIndex        =   23
      Top             =   8280
      Width           =   495
   End
   Begin VB.PictureBox Picture9 
      Height          =   1455
      Left            =   2520
      ScaleHeight     =   1395
      ScaleWidth      =   5955
      TabIndex        =   22
      Top             =   2880
      Visible         =   0   'False
      Width           =   6015
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Next ->"
      Height          =   615
      Left            =   9960
      TabIndex        =   21
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton Command11 
      Caption         =   "<- Previous"
      Height          =   615
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Introduction to Era"
      Height          =   375
      Left            =   4320
      TabIndex        =   19
      Top             =   2280
      Width           =   2295
   End
   Begin VB.PictureBox Picture8 
      Height          =   2655
      Left            =   8400
      ScaleHeight     =   2595
      ScaleWidth      =   2475
      TabIndex        =   18
      Top             =   4800
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.PictureBox Picture6 
      Height          =   3375
      Left            =   2880
      ScaleHeight     =   3315
      ScaleWidth      =   2475
      TabIndex        =   17
      Top             =   5040
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.PictureBox Picture7 
      Height          =   3375
      Left            =   5640
      ScaleHeight     =   3315
      ScaleWidth      =   2595
      TabIndex        =   16
      Top             =   5040
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.PictureBox Picture5 
      Height          =   2655
      Left            =   120
      ScaleHeight     =   2595
      ScaleWidth      =   2595
      TabIndex        =   15
      Top             =   4800
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton Command9 
      Caption         =   "C"
      Height          =   255
      Left            =   10440
      TabIndex        =   14
      Top             =   7560
      Width           =   375
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Info"
      Height          =   255
      Left            =   8640
      TabIndex        =   13
      Top             =   7560
      Width           =   1695
   End
   Begin VB.CommandButton Command7 
      Caption         =   "C"
      Height          =   255
      Left            =   7680
      TabIndex        =   12
      Top             =   8520
      Width           =   375
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Info"
      Height          =   255
      Left            =   5880
      TabIndex        =   11
      Top             =   8520
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "C"
      Height          =   255
      Left            =   4800
      TabIndex        =   10
      Top             =   8520
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Info"
      Height          =   255
      Left            =   3000
      TabIndex        =   9
      Top             =   8520
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "C"
      Height          =   255
      Left            =   2040
      TabIndex        =   8
      Top             =   7560
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Info"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   7560
      Width           =   1695
   End
   Begin VB.PictureBox Picture4 
      Height          =   2655
      Left            =   8400
      Picture         =   "11EarlyRenaissance.frx":32900
      ScaleHeight     =   2595
      ScaleWidth      =   2475
      TabIndex        =   6
      Top             =   4800
      Width           =   2535
   End
   Begin VB.PictureBox Picture3 
      Height          =   3375
      Left            =   5760
      Picture         =   "11EarlyRenaissance.frx":344BF
      ScaleHeight     =   3315
      ScaleWidth      =   2475
      TabIndex        =   5
      Top             =   5040
      Width           =   2535
   End
   Begin VB.PictureBox Picture2 
      Height          =   3375
      Left            =   2880
      Picture         =   "11EarlyRenaissance.frx":366B6
      ScaleHeight     =   3315
      ScaleWidth      =   2475
      TabIndex        =   4
      Top             =   5040
      Width           =   2535
   End
   Begin VB.PictureBox Picture1 
      Height          =   2655
      Left            =   120
      Picture         =   "11EarlyRenaissance.frx":38BB9
      ScaleHeight     =   2595
      ScaleWidth      =   2595
      TabIndex        =   3
      Top             =   4800
      Width           =   2655
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
      Caption         =   "1300 AD ~ 1450 AD"
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
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   1560
      Width           =   11055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Early Renaissance"
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
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   11055
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'WesternArt Project
'Form12
'Bursh,Wrobel
'11-1-06
'This is our Early Renaissance Form Era, displaying works from Era.
Option Explicit
Private Sub Command1_Click()
    Form1.Show
    Form12.Hide
End Sub

Private Sub Command10_Click()
    Picture9.Visible = True
    Picture9.Cls
    Picture9.Print "    With the revival of classical texts, the Renaissance focused on the pursuit of"
    Picture9.Print "humanism, uniting Christianity with Platonic philosophy.  Interests on individual"
    Picture9.Print "fame included territorial, financial, and political power were displayed through "
    Picture9.Print "imagery.  Florence, Italy became the artistic center of the world, as artists gained"
    Picture9.Print "high statue, all due to the well-supported antiquity assimilation."
End Sub

Private Sub Command11_Click()
Form11.Show
Form12.Hide
End Sub

Private Sub Command12_Click()
Form13.Show
Form12.Hide
End Sub

Private Sub Command13_Click()
Form26.Show
Form12.Hide
End Sub

Private Sub Command2_Click()
    Picture1.Visible = False
    Picture5.Visible = True
    Picture5.Cls
    Picture5.Print "The Mérolde Alterpiece"
    Picture5.Print
    Picture5.Print "by Robert Campin"
    Picture5.Print
    Picture5.Print "1425 AD"
End Sub

Private Sub Command3_Click()
    Picture1.Visible = True
    Picture5.Visible = False
    Picture5.Cls
End Sub

Private Sub Command4_Click()
    Picture2.Visible = False
    Picture6.Visible = True
    Picture6.Cls
    Picture6.Print "The Arnolfini Portrait"
    Picture6.Print
    Picture6.Print "by Jan van Eyck"
    Picture6.Print
    Picture6.Print "1434 AD"
End Sub

Private Sub Command5_Click()
    Picture2.Visible = True
    Picture6.Visible = False
    Picture6.Cls
End Sub

Private Sub Command6_Click()
    Picture3.Visible = False
    Picture7.Visible = True
    Picture7.Cls
    Picture7.Print "David"
    Picture7.Print
    Picture7.Print "by Donatello"
    Picture7.Print
    Picture7.Print "1435 AD"
End Sub

Private Sub Command7_Click()
    Picture3.Visible = True
    Picture7.Visible = False
    Picture7.Cls
End Sub

Private Sub Command8_Click()
    Picture4.Visible = False
    Picture8.Visible = True
    Picture8.Cls
    Picture8.Print "Descent from the Cross"
    Picture8.Print
    Picture8.Print "by Rogier van der Weyden"
    Picture8.Print
    Picture8.Print "1435 AD"
End Sub

Private Sub Command9_Click()
    Picture4.Visible = True
    Picture8.Visible = False
    Picture8.Cls
End Sub

Private Sub Picture9_Click()
    Picture9.Visible = False
End Sub
