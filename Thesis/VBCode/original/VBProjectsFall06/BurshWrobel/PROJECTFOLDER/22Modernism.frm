VERSION 5.00
Begin VB.Form Form23 
   Caption         =   "Form23"
   ClientHeight    =   8835
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11085
   LinkTopic       =   "Form23"
   Picture         =   "22Modernism.frx":0000
   ScaleHeight     =   8835
   ScaleWidth      =   11085
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture11 
      Height          =   1335
      Left            =   3000
      ScaleHeight     =   1275
      ScaleWidth      =   4155
      TabIndex        =   26
      Top             =   3240
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Fav"
      Height          =   375
      Left            =   1080
      TabIndex        =   25
      Top             =   8520
      Width           =   495
   End
   Begin VB.CommandButton Command13 
      Caption         =   "<- Previous"
      Height          =   615
      Left            =   0
      TabIndex        =   24
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Introduction to Era"
      Height          =   375
      Left            =   4080
      TabIndex        =   23
      Top             =   2520
      Width           =   2175
   End
   Begin VB.PictureBox Picture10 
      Height          =   1935
      Left            =   7560
      ScaleHeight     =   1875
      ScaleWidth      =   3315
      TabIndex        =   22
      Top             =   3480
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.PictureBox Picture9 
      Height          =   2535
      Left            =   8160
      ScaleHeight     =   2475
      ScaleWidth      =   1995
      TabIndex        =   21
      Top             =   5880
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.PictureBox Picture8 
      Height          =   2175
      Left            =   4440
      ScaleHeight     =   2115
      ScaleWidth      =   2835
      TabIndex        =   20
      Top             =   6000
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.PictureBox Picture7 
      Height          =   2175
      Left            =   1320
      ScaleHeight     =   2115
      ScaleWidth      =   2475
      TabIndex        =   19
      Top             =   6000
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.PictureBox Picture6 
      Height          =   2175
      Left            =   120
      ScaleHeight     =   2115
      ScaleWidth      =   2475
      TabIndex        =   18
      Top             =   3240
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton Command11 
      Caption         =   "C"
      Height          =   255
      Left            =   10080
      TabIndex        =   17
      Top             =   5520
      Width           =   375
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Info"
      Height          =   255
      Left            =   8280
      TabIndex        =   16
      Top             =   5520
      Width           =   1695
   End
   Begin VB.CommandButton Command9 
      Caption         =   "C"
      Height          =   255
      Left            =   9720
      TabIndex        =   15
      Top             =   8520
      Width           =   375
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Info"
      Height          =   255
      Left            =   7920
      TabIndex        =   14
      Top             =   8520
      Width           =   1695
   End
   Begin VB.CommandButton Command7 
      Caption         =   "C"
      Height          =   255
      Left            =   6480
      TabIndex        =   13
      Top             =   8280
      Width           =   375
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Info"
      Height          =   255
      Left            =   4680
      TabIndex        =   12
      Top             =   8280
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "C"
      Height          =   255
      Left            =   3240
      TabIndex        =   11
      Top             =   8280
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Info"
      Height          =   255
      Left            =   1440
      TabIndex        =   10
      Top             =   8280
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "C"
      Height          =   255
      Left            =   2040
      TabIndex        =   9
      Top             =   5520
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Info"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   5520
      Width           =   1695
   End
   Begin VB.PictureBox Picture5 
      Height          =   1935
      Left            =   7560
      Picture         =   "22Modernism.frx":8EF0
      ScaleHeight     =   1875
      ScaleWidth      =   3315
      TabIndex        =   7
      Top             =   3480
      Width           =   3375
   End
   Begin VB.PictureBox Picture4 
      Height          =   2535
      Left            =   8160
      Picture         =   "22Modernism.frx":C487
      ScaleHeight     =   2475
      ScaleWidth      =   1995
      TabIndex        =   6
      Top             =   5880
      Width           =   2055
   End
   Begin VB.PictureBox Picture3 
      Height          =   2175
      Left            =   4440
      Picture         =   "22Modernism.frx":DE17
      ScaleHeight     =   2115
      ScaleWidth      =   2835
      TabIndex        =   5
      Top             =   6000
      Width           =   2895
   End
   Begin VB.PictureBox Picture2 
      Height          =   2175
      Left            =   1320
      Picture         =   "22Modernism.frx":F8DC
      ScaleHeight     =   2115
      ScaleWidth      =   2475
      TabIndex        =   4
      Top             =   6000
      Width           =   2535
   End
   Begin VB.PictureBox Picture1 
      Height          =   2175
      Left            =   120
      Picture         =   "22Modernism.frx":120FF
      ScaleHeight     =   2115
      ScaleWidth      =   2475
      TabIndex        =   3
      Top             =   3240
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Main Menu"
      Height          =   500
      Left            =   0
      TabIndex        =   2
      Top             =   8400
      Width           =   1000
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1960 AD ~ Present"
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
      Height          =   975
      Left            =   240
      TabIndex        =   1
      Top             =   1800
      Width           =   10815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Modernism"
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
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   10815
   End
End
Attribute VB_Name = "Form23"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'WesternArt Project
'Form23
'Bursh,Wrobel
'11-1-06
'This is our Modernism Form Era, displaying works from Era.
Option Explicit
Private Sub Command1_Click()
    Form1.Show
    Form23.Hide
End Sub

Private Sub Command10_Click()
    Picture5.Visible = False
    Picture10.Visible = True
    Picture10.Cls
    Picture10.Print "The Getty Center"
    Picture10.Print
    Picture10.Print "by Richard Meier"
    Picture10.Print
    Picture10.Print "1998 AD"
End Sub

Private Sub Command11_Click()
    Picture5.Visible = True
    Picture10.Visible = False
    Picture10.Cls
End Sub

Private Sub Command12_Click()
    Picture11.Visible = True
    Picture11.Cls
    Picture11.Print "    With age-old themes expressed in new ways, and the"
    Picture11.Print "rapid development of new technology, modern innovation"
    Picture11.Print "seeks to connect all the cultures of the world through"
    Picture11.Print "existing imagery and architecture in a personal or"
    Picture11.Print "original style."
End Sub

Private Sub Command13_Click()
Form22.Show
Form23.Hide
End Sub

Private Sub Command14_Click()
Form26.Show
Form23.Hide
End Sub

Private Sub Command2_Click()
    Picture1.Visible = False
    Picture6.Visible = True
    Picture6.Cls
    Picture6.Print "The Guggenheim Museum"
    Picture6.Print
    Picture6.Print "by Frank Lloyd Wright"
    Picture6.Print
    Picture6.Print "1961 AD"
End Sub

Private Sub Command3_Click()
    Picture1.Visible = True
    Picture6.Visible = False
    Picture6.Cls
End Sub

Private Sub Command4_Click()
    Picture2.Visible = False
    Picture7.Visible = True
    Picture7.Cls
    Picture7.Print "The Whitney Museum"
    Picture7.Print
    Picture7.Print "by Marcel Breuer"
    Picture7.Print
    Picture7.Print "1966 AD"
End Sub

Private Sub Command5_Click()
    Picture2.Visible = True
    Picture7.Visible = False
    Picture7.Cls
End Sub

Private Sub Command6_Click()
    Picture3.Visible = False
    Picture8.Visible = True
    Picture8.Cls
    Picture8.Print "The Spiral Jetty"
    Picture8.Print
    Picture8.Print "by Robert Smithson"
    Picture8.Print
    Picture8.Print "1970 AD"
End Sub

Private Sub Command7_Click()
    Picture3.Visible = True
    Picture8.Visible = False
    Picture8.Cls
End Sub

Private Sub Command8_Click()
    Picture4.Visible = False
    Picture9.Visible = True
    Picture9.Cls
    Picture9.Print "Self-Portrait"
    Picture9.Print
    Picture9.Print "by Chuck Close"
    Picture9.Print
    Picture9.Print "1997 AD"
End Sub


Private Sub Command9_Click()
    Picture4.Visible = True
    Picture9.Visible = False
    Picture9.Cls
End Sub

Private Sub Picture11_Click()
Picture11.Visible = False
End Sub
