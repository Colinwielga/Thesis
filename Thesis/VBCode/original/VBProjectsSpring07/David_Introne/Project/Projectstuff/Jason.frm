VERSION 5.00
Begin VB.Form Jason 
   BackColor       =   &H00808080&
   Caption         =   "Jason"
   ClientHeight    =   7620
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11460
   LinkTopic       =   "Form4"
   ScaleHeight     =   7620
   ScaleWidth      =   11460
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FF80&
      Caption         =   "Next-->"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7695
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00000080&
      Caption         =   "<--Back"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5160
      Width           =   3975
   End
   Begin VB.CommandButton CmdFord 
      BackColor       =   &H0080C0FF&
      Caption         =   "Learn more"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   6255
   End
   Begin VB.PictureBox PicFord 
      BackColor       =   &H80000007&
      FillColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2895
      Left            =   3960
      ScaleHeight     =   2835
      ScaleWidth      =   6195
      TabIndex        =   0
      Top             =   1200
      Width           =   6255
   End
   Begin VB.Image Image2 
      Height          =   3525
      Left            =   3960
      Picture         =   "Jason.frx":0000
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   6195
   End
   Begin VB.Image Image1 
      Height          =   5175
      Left            =   0
      Picture         =   "Jason.frx":11BA9
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3975
   End
End
Attribute VB_Name = "Jason"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CmdFord_Click() ' prints initial message
    PicFord.Cls 'clears then prints Jasons story
    PicFord.Print "Jason: Jason is a younger man"
    PicFord.Print "who owns a small ranch with friends"
    PicFord.Print "in the Colorado Rockies and he"
    PicFord.Print "drives a large SUV. He has an"
    PicFord.Print "affinity for the outdoors. He"
    PicFord.Print "spends most of his time at work"
    PicFord.Print "as the manager of an REI (outdoor"
    PicFord.Print "authority shop) in town. The wheather"
    PicFord.Print "is cold most of the year near by."
End Sub

Private Sub Command1_Click()
If BackToPro < 2 Then 'this goes back so long as you haven't already moved forward as BackToPro counter will show
        Player.Show
        Jason.Hide
    Else
        Select Case puppick 'goes back to a particlur profile base on which puppy you picked
                        Case 11
                            ProShep.Show
                            Jason.Hide
                        Case 12
                            ProPit.Show
                            Jason.Hide
                        Case 13
                            ProMtn.Show
                            Jason.Hide
                        Case 14
                            Produch.Show
                            Jason.Hide
                    End Select
        End If
End Sub

Private Sub Command2_Click()
If BackToPro < 2 Then 'this goes back so long as you haven't already moved forward as BackToPro counter will show
        PupsPick.Show
        Jason.Hide
    Else
        Select Case puppick
                        Case 11 'goes back to a particlur profile base on which puppy you picked
                            ProShep.Show
                            Jason.Hide
                        Case 12
                            ProPit.Show
                            Jason.Hide
                        Case 13
                            ProMtn.Show
                            Jason.Hide
                        Case 14
                            Produch.Show
                            Jason.Hide
                    End Select
    
    End If
End Sub

