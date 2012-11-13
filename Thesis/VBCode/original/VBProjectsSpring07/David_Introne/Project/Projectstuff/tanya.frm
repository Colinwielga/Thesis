VERSION 5.00
Begin VB.Form tanya 
   BackColor       =   &H00008000&
   Caption         =   "Tanya"
   ClientHeight    =   8160
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11130
   LinkTopic       =   "Form1"
   ScaleHeight     =   8160
   ScaleWidth      =   11130
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFC0FF&
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
      Height          =   735
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7440
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0FF&
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
      Height          =   735
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7440
      Width           =   1695
   End
   Begin VB.PictureBox Pictanya 
      BackColor       =   &H00004000&
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   2895
      Left            =   2160
      ScaleHeight     =   2835
      ScaleWidth      =   7035
      TabIndex        =   1
      Top             =   5280
      Width           =   7095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Learn More"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   6255
   End
   Begin VB.Image Image2 
      Height          =   3525
      Left            =   5160
      Picture         =   "tanya.frx":0000
      Top             =   1320
      Width           =   5715
   End
   Begin VB.Image Image1 
      Height          =   5295
      Left            =   0
      Picture         =   "tanya.frx":19E3F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4935
   End
End
Attribute VB_Name = "tanya"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
    Pictanya.Cls
    Pictanya.Print "Tanya is teacher, who drives"
    Pictanya.Print "a Lexus, and owns her home in"
    Pictanya.Print "Maryland. The backyard has a large"
    Pictanya.Print "river running 45 yards behind it."
    Pictanya.Print "She has a family, and two children."
    Pictanya.Print "She lives near her school, but far"
    Pictanya.Print "from any parks"
End Sub

Private Sub Command2_Click()
    If BackToPro < 2 Then
                Player.Show
                tanya.Hide
            Else
                Select Case puppick
                            Case 11
                                ProShep.Show
                                tanya.Hide
                            Case 12
                                ProPit.Show
                                tanya.Hide
                            Case 13
                                ProMtn.Show
                                tanya.Hide
                            Case 14
                                Produch.Show
                                tanya.Hide
                End Select
    End If
End Sub

Private Sub Command3_Click()
 If BackToPro < 2 Then
                PupsPick.Show
                tanya.Hide
            Else
                Select Case puppick
                                Case 11
                                    ProShep.Show
                                    tanya.Hide
                                Case 12
                                    ProPit.Show
                                    tanya.Hide
                                Case 13
                                    ProMtn.Show
                                    tanya.Hide
                                Case 14
                                    Produch.Show
                                    tanya.Hide
                            End Select
    End If
End Sub
