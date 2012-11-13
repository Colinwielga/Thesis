VERSION 5.00
Begin VB.Form Form20 
   BackColor       =   &H00008000&
   Caption         =   "Form20"
   ClientHeight    =   8835
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11085
   LinkTopic       =   "Form20"
   Picture         =   "19Imperssionism.frx":0000
   ScaleHeight     =   8835
   ScaleWidth      =   11085
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command15 
      Caption         =   "Fav"
      Height          =   375
      Left            =   1080
      TabIndex        =   27
      Top             =   8400
      Width           =   495
   End
   Begin VB.PictureBox Picture11 
      Height          =   1455
      Left            =   3360
      ScaleHeight     =   1395
      ScaleWidth      =   4035
      TabIndex        =   26
      Top             =   3120
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Next ->"
      Height          =   615
      Left            =   9840
      TabIndex        =   25
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton Command13 
      Caption         =   "<- Previous"
      Height          =   615
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Introduction to Era"
      Height          =   375
      Left            =   4200
      TabIndex        =   23
      Top             =   2400
      Width           =   2295
   End
   Begin VB.PictureBox Picture10 
      Height          =   2055
      Left            =   8760
      ScaleHeight     =   1995
      ScaleWidth      =   2115
      TabIndex        =   22
      Top             =   6480
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.PictureBox Picture9 
      Height          =   2415
      Left            =   8280
      ScaleHeight     =   2355
      ScaleWidth      =   2355
      TabIndex        =   21
      Top             =   3480
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.PictureBox Picture8 
      Height          =   2535
      Left            =   4800
      ScaleHeight     =   2475
      ScaleWidth      =   2955
      TabIndex        =   20
      Top             =   5520
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.PictureBox Picture7 
      Height          =   3375
      Left            =   2160
      ScaleHeight     =   3315
      ScaleWidth      =   2235
      TabIndex        =   19
      Top             =   5160
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.PictureBox Picture6 
      Height          =   1935
      Left            =   0
      ScaleHeight     =   1875
      ScaleWidth      =   2835
      TabIndex        =   18
      Top             =   2880
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton Command11 
      Caption         =   "C"
      Height          =   255
      Left            =   10560
      TabIndex        =   17
      Top             =   8640
      Width           =   375
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Info"
      Height          =   255
      Left            =   8760
      TabIndex        =   16
      Top             =   8640
      Width           =   1695
   End
   Begin VB.CommandButton Command9 
      Caption         =   "C"
      Height          =   255
      Left            =   10200
      TabIndex        =   15
      Top             =   6000
      Width           =   375
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Info"
      Height          =   255
      Left            =   8400
      TabIndex        =   14
      Top             =   6000
      Width           =   1695
   End
   Begin VB.CommandButton Command7 
      Caption         =   "C"
      Height          =   255
      Left            =   6960
      TabIndex        =   13
      Top             =   8160
      Width           =   375
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Info"
      Height          =   255
      Left            =   5160
      TabIndex        =   12
      Top             =   8160
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "C"
      Height          =   255
      Left            =   3840
      TabIndex        =   11
      Top             =   8640
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Info"
      Height          =   255
      Left            =   2160
      TabIndex        =   10
      Top             =   8640
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "C"
      Height          =   255
      Left            =   2040
      TabIndex        =   9
      Top             =   4920
      Width           =   385
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Info"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   4920
      Width           =   1695
   End
   Begin VB.PictureBox Picture5 
      Height          =   2055
      Left            =   8760
      Picture         =   "19Imperssionism.frx":12896
      ScaleHeight     =   1995
      ScaleWidth      =   2115
      TabIndex        =   7
      Top             =   6480
      Width           =   2175
   End
   Begin VB.PictureBox Picture4 
      Height          =   2415
      Left            =   8280
      Picture         =   "19Imperssionism.frx":147D9
      ScaleHeight     =   2355
      ScaleWidth      =   2355
      TabIndex        =   6
      Top             =   3480
      Width           =   2415
   End
   Begin VB.PictureBox Picture3 
      Height          =   2535
      Left            =   4800
      Picture         =   "19Imperssionism.frx":16EDE
      ScaleHeight     =   2475
      ScaleWidth      =   2955
      TabIndex        =   5
      Top             =   5520
      Width           =   3015
   End
   Begin VB.PictureBox Picture2 
      Height          =   3375
      Left            =   2160
      Picture         =   "19Imperssionism.frx":19EC0
      ScaleHeight     =   3315
      ScaleWidth      =   2235
      TabIndex        =   4
      Top             =   5160
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      Height          =   1935
      Left            =   0
      Picture         =   "19Imperssionism.frx":1C2C2
      ScaleHeight     =   1875
      ScaleWidth      =   2835
      TabIndex        =   3
      Top             =   2880
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Main Menu"
      Height          =   500
      Left            =   0
      TabIndex        =   2
      Top             =   8280
      Width           =   1000
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1870 AD ~ 1900 AD"
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
      Top             =   1800
      Width           =   11055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Impressionism"
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
      Top             =   120
      Width           =   11055
   End
End
Attribute VB_Name = "Form20"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'WesternArt Project
'Form20
'Bursh,Wrobel
'11-1-06
'This is our Impressionism Form Era, displaying works from Era.
Option Explicit
Private Sub Command1_Click()
    Form1.Show
    Form20.Hide
End Sub

Private Sub Command10_Click()
    Picture5.Visible = False
    Picture10.Visible = True
    Picture10.Cls
    Picture10.Print "Water Lily Pond"
    Picture10.Print
    Picture10.Print "by Claude Monet"
    Picture10.Print
    Picture10.Print "1899 AD"

End Sub

Private Sub Command11_Click()
    Picture5.Visible = True
    Picture10.Visible = False
    Picture10.Cls
End Sub

Private Sub Command12_Click()
    Picture11.Visible = True
    Picture11.Cls
    Picture11.Print "    Rarely responding to political events - Impressionism"
    Picture11.Print "art consisted of leisure activities, entertainment, within"
    Picture11.Print "landscapes and cityscapes, with focus on changes in"
    Picture11.Print "lightand color as time of day passes, in a non-tedious,"
    Picture11.Print "freestroking style labeled 'Art for Art's Sake.'"
End Sub

Private Sub Command13_Click()
Form19.Show
Form20.Hide
End Sub

Private Sub Command14_Click()
Form21.Show
Form20.Hide
End Sub

Private Sub Command15_Click()
    Form26.Show
    Form20.Hide
End Sub

Private Sub Command2_Click()
    Picture1.Visible = False
    Picture6.Visible = True
    Picture6.Cls
    Picture6.Print "Breezing Up (A Fair Wind)"
    Picture6.Print
    Picture6.Print "by Winslow Homer"
    Picture6.Print
    Picture6.Print "1875 AD"
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
    Picture7.Print "Nocturne in Black and Gold"
    Picture7.Print
    Picture7.Print "by James Abbott McNeill"
    Picture7.Print "Whistler"
    Picture7.Print
    Picture7.Print "1875 AD"

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
    Picture8.Print "Moulin de la Galette"
    Picture8.Print
    Picture8.Print "by Pierre-Auguste Renoir"
    Picture8.Print
    Picture8.Print "1876 AD"
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
    Picture9.Print "Place du Théâtre Français"
    Picture9.Print
    Picture9.Print "by Camille Pissarro"
    Picture9.Print
    Picture9.Print "1898 AD"


End Sub

Private Sub Command9_Click()
    Picture4.Visible = True
    Picture9.Visible = False
    Picture9.Cls
End Sub

Private Sub Picture11_Click()
Picture11.Visible = False
End Sub
