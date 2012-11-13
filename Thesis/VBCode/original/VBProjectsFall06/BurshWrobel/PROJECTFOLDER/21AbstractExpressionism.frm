VERSION 5.00
Begin VB.Form Form22 
   Caption         =   "Form22"
   ClientHeight    =   8835
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11085
   LinkTopic       =   "Form22"
   Picture         =   "21AbstractExpressionism.frx":0000
   ScaleHeight     =   8835
   ScaleWidth      =   11085
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command13 
      Caption         =   "Fav"
      Height          =   375
      Left            =   1080
      TabIndex        =   23
      Top             =   8520
      Width           =   495
   End
   Begin VB.PictureBox Picture9 
      Height          =   615
      Left            =   840
      ScaleHeight     =   555
      ScaleWidth      =   9675
      TabIndex        =   22
      Top             =   4080
      Visible         =   0   'False
      Width           =   9735
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Next ->"
      Height          =   615
      Left            =   9840
      TabIndex        =   21
      Top             =   120
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
      Left            =   4440
      TabIndex        =   19
      Top             =   3600
      Width           =   2295
   End
   Begin VB.PictureBox Picture8 
      Height          =   3135
      Left            =   8400
      ScaleHeight     =   3075
      ScaleWidth      =   2115
      TabIndex        =   18
      Top             =   4800
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.PictureBox Picture7 
      Height          =   3135
      Left            =   5760
      ScaleHeight     =   3075
      ScaleWidth      =   2115
      TabIndex        =   17
      Top             =   4800
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.PictureBox Picture6 
      Height          =   3135
      Left            =   3240
      ScaleHeight     =   3075
      ScaleWidth      =   2115
      TabIndex        =   16
      Top             =   4800
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.PictureBox Picture5 
      Height          =   3135
      Left            =   600
      ScaleHeight     =   3075
      ScaleWidth      =   2115
      TabIndex        =   15
      Top             =   4800
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton Command9 
      Caption         =   "C"
      Height          =   255
      Left            =   10200
      TabIndex        =   14
      Top             =   8040
      Width           =   375
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Info"
      Height          =   255
      Left            =   8400
      TabIndex        =   13
      Top             =   8040
      Width           =   1695
   End
   Begin VB.CommandButton Command7 
      Caption         =   "C"
      Height          =   255
      Left            =   7560
      TabIndex        =   12
      Top             =   8040
      Width           =   375
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Info"
      Height          =   255
      Left            =   5760
      TabIndex        =   11
      Top             =   8040
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "C"
      Height          =   255
      Left            =   5040
      TabIndex        =   10
      Top             =   8040
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Info"
      Height          =   255
      Left            =   3240
      TabIndex        =   9
      Top             =   8040
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "C"
      Height          =   255
      Left            =   2400
      TabIndex        =   8
      Top             =   8040
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Info"
      Height          =   255
      Left            =   600
      TabIndex        =   7
      Top             =   8040
      Width           =   1695
   End
   Begin VB.PictureBox Picture4 
      Height          =   3135
      Left            =   8400
      Picture         =   "21AbstractExpressionism.frx":165D1
      ScaleHeight     =   3075
      ScaleWidth      =   2115
      TabIndex        =   6
      Top             =   4800
      Width           =   2175
   End
   Begin VB.PictureBox Picture3 
      Height          =   3135
      Left            =   5760
      Picture         =   "21AbstractExpressionism.frx":19FC0
      ScaleHeight     =   3075
      ScaleWidth      =   2115
      TabIndex        =   5
      Top             =   4800
      Width           =   2175
   End
   Begin VB.PictureBox Picture2 
      Height          =   3135
      Left            =   3240
      Picture         =   "21AbstractExpressionism.frx":1EF68
      ScaleHeight     =   3075
      ScaleWidth      =   2115
      TabIndex        =   4
      Top             =   4800
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      Height          =   3135
      Left            =   600
      Picture         =   "21AbstractExpressionism.frx":24498
      ScaleHeight     =   3075
      ScaleWidth      =   2115
      TabIndex        =   3
      Top             =   4800
      Width           =   2175
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
      Caption         =   "1950 AD ~ 1960 AD"
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
      Left            =   120
      TabIndex        =   1
      Top             =   3000
      Width           =   10935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Abstract Expressionism"
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
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   11055
   End
End
Attribute VB_Name = "Form22"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'WesternArt Project
'Form22
'Bursh,Wrobel
'11-1-06
'This is our Abstract Expressionism Form Era, displaying works from Era.
Option Explicit
Private Sub Command1_Click()
    Form1.Show
    Form22.Hide
End Sub

Private Sub Command10_Click()
    Picture9.Visible = True
    Picture9.Cls
    Picture9.Print "    When Paris had fallen to Nazi Germany, the center of the art world shifted to New York.  Abstract Expressionism deals with the new"
    Picture9.Print "original ways of the artmaking process, the majority of which are non-objective.  Contrasting colors, shapes, lines in a design unique"
    Picture9.Print "to the individual artist was their way of emotionally expressing themselves."
End Sub

Private Sub Command11_Click()
Form21.Show
Form22.Hide
End Sub

Private Sub Command12_Click()
Form23.Show
Form22.Hide
End Sub

Private Sub Command13_Click()
Form26.Show
Form22.Hide
End Sub

Private Sub Command2_Click()
    Picture1.Visible = False
    Picture5.Visible = True
    Picture5.Cls
    Picture5.Print "Woman and Bicycle"
    Picture5.Print
    Picture5.Print "by Willem de Koonig"
    Picture5.Print
    Picture5.Print "1952 AD"
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
    Picture6.Print "White Light"
    Picture6.Print
    Picture6.Print "by Jackson Pollock"
    Picture6.Print
    Picture6.Print "1954 AD"
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
    Picture7.Print "No. 15"
    Picture7.Print
    Picture7.Print "by Mark Rothko"
    Picture7.Print
    Picture7.Print "1957 AD"
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
    Picture8.Print "Black Wall"
    Picture8.Print
    Picture8.Print "by Louise Nevelson"
    Picture8.Print
    Picture8.Print "1959 AD"
End Sub

Private Sub Command9_Click()
    Picture4.Visible = True
    Picture8.Visible = False
    Picture8.Cls
End Sub

Private Sub Picture9_Click()
    Picture9.Visible = False
End Sub
