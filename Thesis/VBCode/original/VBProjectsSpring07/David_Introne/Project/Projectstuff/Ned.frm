VERSION 5.00
Begin VB.Form Ned 
   BackColor       =   &H00808000&
   Caption         =   "Ned"
   ClientHeight    =   8655
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13320
   FillColor       =   &H000040C0&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8655
   ScaleWidth      =   13320
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Next-->"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7320
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "<--Back"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7200
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "Learn More"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4320
      Width           =   2415
   End
   Begin VB.PictureBox PicResult1 
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   3240
      ScaleHeight     =   2835
      ScaleWidth      =   6075
      TabIndex        =   0
      Top             =   5760
      Width           =   6135
   End
   Begin VB.Image Image2 
      Height          =   6840
      Left            =   3240
      Picture         =   "Ned.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10080
   End
   Begin VB.Image Image1 
      Height          =   4860
      Left            =   0
      Picture         =   "Ned.frx":5FE01
      Top             =   0
      Width           =   3465
   End
End
Attribute VB_Name = "Ned"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
    PicResult1.Cls 'clears then prints the story
    PicResult1.Print " Ned: Ned is fun loving and a generally"
    PicResult1.Print " cheerful fellow. Located in north"
    PicResult1.Print " Minneapolis, his house is small,"
    PicResult1.Print " with a one car garage, and modest"
    PicResult1.Print " backyard, which he keeps well trimmed"
    PicResult1.Print " and looking sharp. However Ned"
    PicResult1.Print " has hidden charm, he Drives a"
    PicResult1.Print " 1961 Ferrari 250 GT California."
End Sub

Private Sub Command2_Click()
    If BackToPro < 2 Then
            Player.Show 'this goes back so long as you haven't already moved forward as BackToPro counter will show
            Ned.Hide
        Else
            Select Case puppick 'goes back to a particlur profile base on which puppy you picked
                Case 11
                    ProShep.Show
                    Ned.Hide
                Case 12
                    ProPit.Show
                    Ned.Hide
                Case 13
                    ProMtn.Show
                    Ned.Hide
                Case 14
                    Produch.Show
                    Ned.Hide
            End Select
    End If
End Sub

Private Sub Command3_Click()
    If BackToPro < 2 Then
            PupsPick.Show
            Ned.Hide
        Else
            Select Case puppick
                            Case 11
                                ProShep.Show
                                Ned.Hide
                            Case 12
                                ProPit.Show
                                Ned.Hide
                            Case 13
                                ProMtn.Show
                                Ned.Hide
                            Case 14
                                Produch.Show
                                Ned.Hide
                        End Select
    End If
End Sub
