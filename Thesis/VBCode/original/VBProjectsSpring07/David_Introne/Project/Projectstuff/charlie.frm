VERSION 5.00
Begin VB.Form charlie 
   BackColor       =   &H0000FFFF&
   Caption         =   "Charlie"
   ClientHeight    =   8415
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11640
   LinkTopic       =   "Form3"
   ScaleHeight     =   8415
   ScaleWidth      =   11640
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      BackColor       =   &H008080FF&
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
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7440
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
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
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7440
      Width           =   1815
   End
   Begin VB.PictureBox picsly 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   4560
      ScaleHeight     =   3675
      ScaleWidth      =   6555
      TabIndex        =   1
      Top             =   1680
      Width           =   6615
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "Learn More"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
   Begin VB.Image Image2 
      Height          =   2610
      Left            =   6360
      Picture         =   "charlie.frx":0000
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   3480
   End
   Begin VB.Image Image1 
      Height          =   2775
      Left            =   1800
      Picture         =   "charlie.frx":2A73
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   4575
   End
   Begin VB.Image charliepic 
      Height          =   5415
      Left            =   0
      Picture         =   "charlie.frx":56E3
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4095
   End
End
Attribute VB_Name = "charlie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
    picsly.Cls 'clears then prints the story
    picsly.Print "Charlie: Charlie lives in Philadelphia,"
    picsly.Print "where he works at an olive oil company"
    picsly.Print "for his cousin's, brother's friend who"
    picsly.Print "knows a guy. Charlie doesn't drive."
    picsly.Print "The neighborhood is questionable and"
    picsly.Print "he lives in tiny apartment with no"
    picsly.Print "backyard. However he does seem to "
    picsly.Print "enjoy a good workout. Oh and he's a"
    picsly.Print "felon."
End Sub

Private Sub Command2_Click()
If BackToPro < 2 Then
        PupsPick.Show 'this goes back so long as you haven't already moved forward as BackToPro counter will show
        charlie.Hide
    Else
        Select Case puppick 'goes back to a particlur profile base on which puppy you picked
                        Case 11
                            ProShep.Show
                            charlie.Hide
                        Case 12
                            ProPit.Show
                            charlie.Hide
                        Case 13
                            ProMtn.Show
                            charlie.Hide
                        Case 14
                            Produch.Show
                            charlie.Hide
                    End Select
    
    End If
End Sub

Private Sub Command3_Click()
 If BackToPro < 2 Then
        charlie.Hide
        Player.Show 'goes back to a particlur profile base on which puppy you picked
    Else
        Select Case puppick
                        Case 11
                            ProShep.Show 'goes back to a particlur profile base on which puppy you picked
                            charlie.Hide
                        Case 12
                            ProPit.Show
                            charlie.Hide
                        Case 13
                            ProMtn.Show
                            charlie.Hide
                        Case 14
                            Produch.Show
                            charlie.Hide
                    End Select
    
    End If
End Sub


