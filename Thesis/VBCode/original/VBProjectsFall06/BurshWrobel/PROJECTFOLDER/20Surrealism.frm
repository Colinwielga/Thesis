VERSION 5.00
Begin VB.Form Form21 
   Caption         =   "Form21"
   ClientHeight    =   8835
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11085
   LinkTopic       =   "Form21"
   Picture         =   "20Surrealism.frx":0000
   ScaleHeight     =   8835
   ScaleWidth      =   11085
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton Command23 
      Caption         =   "Fav"
      Height          =   375
      Left            =   120
      TabIndex        =   43
      Top             =   7800
      Width           =   495
   End
   Begin VB.PictureBox Picture19 
      Height          =   1455
      Left            =   2760
      ScaleHeight     =   1395
      ScaleWidth      =   5715
      TabIndex        =   42
      Top             =   2640
      Visible         =   0   'False
      Width           =   5775
   End
   Begin VB.CommandButton Command22 
      Caption         =   "Next ->"
      Height          =   615
      Left            =   9840
      TabIndex        =   41
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton Command21 
      Caption         =   "<- Previous"
      Height          =   615
      Left            =   0
      TabIndex        =   40
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton Command20 
      Caption         =   "Introduction to Era"
      Height          =   375
      Left            =   4320
      TabIndex        =   39
      Top             =   2160
      Width           =   2295
   End
   Begin VB.PictureBox Picture18 
      Height          =   2295
      Left            =   9000
      ScaleHeight     =   2235
      ScaleWidth      =   1635
      TabIndex        =   38
      Top             =   1680
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.PictureBox Picture17 
      Height          =   1815
      Left            =   7560
      ScaleHeight     =   1755
      ScaleWidth      =   3075
      TabIndex        =   37
      Top             =   4320
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.PictureBox Picture16 
      Height          =   2175
      Left            =   8640
      ScaleHeight     =   2115
      ScaleWidth      =   1995
      TabIndex        =   36
      Top             =   6360
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.PictureBox Picture15 
      Height          =   2175
      Left            =   6240
      ScaleHeight     =   2115
      ScaleWidth      =   1635
      TabIndex        =   35
      Top             =   6360
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.PictureBox Picture14 
      Height          =   2175
      Left            =   4320
      ScaleHeight     =   2115
      ScaleWidth      =   1635
      TabIndex        =   34
      Top             =   6360
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.PictureBox Picture13 
      Height          =   1455
      Left            =   1320
      ScaleHeight     =   1395
      ScaleWidth      =   1875
      TabIndex        =   33
      Top             =   6840
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.PictureBox Picture12 
      Height          =   2295
      Left            =   2520
      ScaleHeight     =   2235
      ScaleWidth      =   1635
      TabIndex        =   32
      Top             =   4200
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.PictureBox Picture11 
      Height          =   2295
      Left            =   240
      ScaleHeight     =   2235
      ScaleWidth      =   1635
      TabIndex        =   31
      Top             =   4200
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.PictureBox Picture10 
      Height          =   1815
      Left            =   0
      ScaleHeight     =   1755
      ScaleWidth      =   2235
      TabIndex        =   30
      Top             =   2040
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton Command19 
      Caption         =   "C"
      Height          =   255
      Left            =   10560
      TabIndex        =   29
      Top             =   3960
      Width           =   375
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Info"
      Height          =   255
      Left            =   8880
      TabIndex        =   28
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CommandButton Command17 
      Caption         =   "C"
      Height          =   255
      Left            =   9720
      TabIndex        =   27
      Top             =   6120
      Width           =   375
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Info"
      Height          =   255
      Left            =   8040
      TabIndex        =   26
      Top             =   6120
      Width           =   1695
   End
   Begin VB.CommandButton Command15 
      Caption         =   "C"
      Height          =   255
      Left            =   10320
      TabIndex        =   25
      Top             =   8520
      Width           =   375
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Info"
      Height          =   255
      Left            =   8640
      TabIndex        =   24
      Top             =   8520
      Width           =   1695
   End
   Begin VB.CommandButton Command13 
      Caption         =   "C"
      Height          =   255
      Left            =   7800
      TabIndex        =   23
      Top             =   8520
      Width           =   375
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Info"
      Height          =   255
      Left            =   6120
      TabIndex        =   22
      Top             =   8520
      Width           =   1695
   End
   Begin VB.CommandButton Command11 
      Caption         =   "C"
      Height          =   255
      Left            =   5640
      TabIndex        =   21
      Top             =   8520
      Width           =   375
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Info"
      Height          =   255
      Left            =   3960
      TabIndex        =   20
      Top             =   8520
      Width           =   1695
   End
   Begin VB.CommandButton Command9 
      Caption         =   "C"
      Height          =   255
      Left            =   3000
      TabIndex        =   19
      Top             =   8400
      Width           =   375
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Info"
      Height          =   255
      Left            =   1200
      TabIndex        =   18
      Top             =   8400
      Width           =   1695
   End
   Begin VB.CommandButton Command7 
      Caption         =   "C"
      Height          =   255
      Left            =   3960
      TabIndex        =   17
      Top             =   6480
      Width           =   375
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Info"
      Height          =   255
      Left            =   2280
      TabIndex        =   16
      Top             =   6480
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "C"
      Height          =   255
      Left            =   1680
      TabIndex        =   15
      Top             =   6480
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Info"
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   6480
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "C"
      Height          =   255
      Left            =   1800
      TabIndex        =   13
      Top             =   3840
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Info"
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   3840
      Width           =   1695
   End
   Begin VB.PictureBox Picture6 
      Height          =   2295
      Left            =   6240
      Picture         =   "20Surrealism.frx":168B8
      ScaleHeight     =   2235
      ScaleWidth      =   1635
      TabIndex        =   11
      Top             =   6360
      Width           =   1695
   End
   Begin VB.PictureBox Picture9 
      Height          =   2295
      Left            =   9000
      Picture         =   "20Surrealism.frx":17772
      ScaleHeight     =   2235
      ScaleWidth      =   1635
      TabIndex        =   10
      Top             =   1680
      Width           =   1695
   End
   Begin VB.PictureBox Picture8 
      Height          =   1815
      Left            =   7560
      Picture         =   "20Surrealism.frx":1890B
      ScaleHeight     =   1755
      ScaleWidth      =   3075
      TabIndex        =   9
      Top             =   4320
      Width           =   3135
   End
   Begin VB.PictureBox Picture7 
      Height          =   2295
      Left            =   8640
      Picture         =   "20Surrealism.frx":1A8C6
      ScaleHeight     =   2235
      ScaleWidth      =   1995
      TabIndex        =   8
      Top             =   6360
      Width           =   2055
   End
   Begin VB.PictureBox Picture5 
      Height          =   2295
      Left            =   4320
      Picture         =   "20Surrealism.frx":1CCDA
      ScaleHeight     =   2235
      ScaleWidth      =   1635
      TabIndex        =   7
      Top             =   6360
      Width           =   1695
   End
   Begin VB.PictureBox Picture4 
      Height          =   1455
      Left            =   1320
      Picture         =   "20Surrealism.frx":1DD30
      ScaleHeight     =   1395
      ScaleWidth      =   1875
      TabIndex        =   6
      Top             =   6840
      Width           =   1935
   End
   Begin VB.PictureBox Picture3 
      Height          =   2295
      Left            =   2520
      Picture         =   "20Surrealism.frx":1F272
      ScaleHeight     =   2235
      ScaleWidth      =   1635
      TabIndex        =   5
      Top             =   4200
      Width           =   1695
   End
   Begin VB.PictureBox Picture2 
      Height          =   2295
      Left            =   240
      Picture         =   "20Surrealism.frx":201B0
      ScaleHeight     =   2235
      ScaleWidth      =   1635
      TabIndex        =   4
      Top             =   4200
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      Height          =   1815
      Left            =   0
      Picture         =   "20Surrealism.frx":215CC
      ScaleHeight     =   1755
      ScaleWidth      =   2235
      TabIndex        =   3
      Top             =   2040
      Width           =   2295
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
      Caption         =   "1900 AD ~ 1950 AD"
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
      Caption         =   "Surrealism "
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
Attribute VB_Name = "Form21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'WesternArt Project
'Form21
'Bursh,Wrobel
'11-1-06
'This is our Surrealism Form Era, displaying works from Era.
Option Explicit
Private Sub Command1_Click()
    Form1.Show
    Form21.Hide
End Sub

Private Sub Command10_Click()
    Picture5.Visible = False
    Picture14.Visible = True
    Picture14.Cls
    Picture14.Print "L.H.O.O.Q."
    Picture14.Print
    Picture14.Print "by Marcel Duchamp"
    Picture14.Print
    Picture14.Print "1915 AD"
End Sub

Private Sub Command11_Click()
    Picture5.Visible = True
    Picture14.Visible = False
    Picture14.Cls
End Sub

Private Sub Command12_Click()
    Picture6.Visible = False
    Picture15.Visible = True
    Picture15.Cls
    Picture15.Print "Fountain"
    Picture15.Print
    Picture15.Print "by Marcel Duchamp"
    Picture15.Print
    Picture15.Print "1917 AD"
End Sub

Private Sub Command13_Click()
    Picture6.Visible = True
    Picture15.Visible = False
    Picture15.Cls
End Sub

Private Sub Command14_Click()
   Picture7.Visible = False
    Picture16.Visible = True
    Picture16.Cls
    Picture16.Print "American Gothic"
    Picture16.Print
    Picture16.Print "by Grant Wood"
    Picture16.Print
    Picture16.Print "1930 AD"
End Sub

Private Sub Command15_Click()
    Picture7.Visible = True
    Picture16.Visible = False
    Picture16.Cls
End Sub

Private Sub Command16_Click()
    Picture8.Visible = False
    Picture17.Visible = True
    Picture17.Cls
    Picture17.Print "Guernica"
    Picture17.Print
    Picture17.Print "by Pablo Picasso"
    Picture17.Print
    Picture17.Print "1937 AD"
End Sub

Private Sub Command17_Click()
    Picture8.Visible = True
    Picture17.Visible = False
    Picture17.Cls
End Sub

Private Sub Command18_Click()
    Picture9.Visible = False
    Picture18.Visible = True
    Picture18.Cls
    Picture18.Print "Time Transfixed"
    Picture18.Print
    Picture18.Print "by Rene Magritte"
    Picture18.Print
    Picture18.Print "1938 AD"
End Sub

Private Sub Command19_Click()
    Picture9.Visible = True
    Picture18.Visible = False
    Picture18.Cls
End Sub

Private Sub Command2_Click()
    Picture1.Visible = False
    Picture10.Visible = True
    Picture10.Cls
    Picture10.Print "The Starry Night"
    Picture10.Print
    Picture10.Print "by Vincent Van Gogh"
    Picture10.Print
    Picture10.Print "1900 AD"
End Sub

Private Sub Command20_Click()
Picture19.Visible = True
    Picture19.Cls
    Picture19.Print "    Surrealism comes from a wide variety of movements in the early 20th century"
    Picture19.Print "many a result of World War I.  This period redefined the meaning of art, with a"
    Picture19.Print "'newness' quality, emotional content, and symbolism, while using vibrant changes"
    Picture19.Print "color and light, dreaming, imagination, and rebellious imagery with the content of"
    Picture19.Print "starting life over."
End Sub

Private Sub Command21_Click()
Form20.Show
Form21.Hide
End Sub

Private Sub Command22_Click()
Form22.Show
Form21.Hide
End Sub

Private Sub Command23_Click()
    Form26.Show
    Form21.Hide
End Sub

Private Sub Command3_Click()
    Picture1.Visible = True
    Picture10.Visible = False
    Picture10.Cls
End Sub

Private Sub Command4_Click()
    Picture2.Visible = False
    Picture11.Visible = True
    Picture11.Cls
    Picture11.Print "The Scream"
    Picture11.Print
    Picture11.Print "by Edvard Munch"
    Picture11.Print
    Picture11.Print "1900 AD"
End Sub

Private Sub Command5_Click()
    Picture2.Visible = True
    Picture11.Visible = False
    Picture11.Cls
End Sub

Private Sub Command6_Click()
    Picture3.Visible = False
    Picture12.Visible = True
    Picture12.Cls
    Picture12.Print "The Old Guitarist"
    Picture12.Print
    Picture12.Print "by Pablo Picasso"
    Picture12.Print
    Picture12.Print "1903 AD"
End Sub

Private Sub Command7_Click()
    Picture3.Visible = True
    Picture12.Visible = False
    Picture12.Cls
End Sub

Private Sub Command8_Click()
    Picture4.Visible = False
    Picture13.Visible = True
    Picture13.Cls
    Picture13.Print "Harmony in Red"
    Picture13.Print
    Picture13.Print "by Henri Matisse"
    Picture13.Print
    Picture13.Print "1908 AD"
End Sub

Private Sub Command9_Click()
    Picture4.Visible = True
    Picture13.Visible = False
    Picture13.Cls
End Sub

Private Sub Picture19_Click()
    Picture19.Visible = False
End Sub
