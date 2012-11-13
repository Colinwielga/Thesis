VERSION 5.00
Begin VB.Form sheryl 
   BackColor       =   &H000000C0&
   Caption         =   "Sheryl"
   ClientHeight    =   7695
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10530
   LinkTopic       =   "Form2"
   ScaleHeight     =   7695
   ScaleWidth      =   10530
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFC0&
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
      Height          =   855
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6840
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Next -->"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6840
      Width           =   1695
   End
   Begin VB.PictureBox PicSheryl 
      BackColor       =   &H00000040&
      FillColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   4215
      Left            =   0
      ScaleHeight     =   4155
      ScaleWidth      =   6675
      TabIndex        =   1
      Top             =   0
      Width           =   6735
   End
   Begin VB.CommandButton CmdSheryrl 
      BackColor       =   &H00FFFF80&
      Caption         =   "Learn more"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6000
      Width           =   1815
   End
   Begin VB.Image Image2 
      Height          =   3525
      Left            =   -360
      Picture         =   "sheryl.frx":0000
      Top             =   4200
      Width           =   5715
   End
   Begin VB.Image Image1 
      Height          =   5580
      Left            =   6720
      Picture         =   "sheryl.frx":180B5
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3780
   End
End
Attribute VB_Name = "sheryl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CmdSheryrl_Click()
    PicSheryl.Cls
    PicSheryl.Print "Sheryl: Sheryl lives in New York,"
    PicSheryl.Print "her apartment overlooks the Park."
    PicSheryl.Print "Her car is a VW Jetta and she loves"
    PicSheryl.Print "to run. Unfortunately she has a"
    PicSheryl.Print "very small back balcony and she"
    PicSheryl.Print "also hates kids."
End Sub

Private Sub Command1_Click()
    If BackToPro < 2 Then
                PupsPick.Show
                sheryl.Hide
            Else
                Select Case puppick
                        Case 11
                            ProShep.Show
                            sheryl.Hide
                        Case 12
                            ProPit.Show
                            sheryl.Hide
                        Case 13
                            ProMtn.Show
                            sheryl.Hide
                        Case 14
                            Produch.Show
                            sheryl.Hide
                End Select
    End If
End Sub

Private Sub Command2_Click()
    If BackToPro < 2 Then
                Player.Show
                sheryl.Hide
            Else
                Select Case puppick
                    Case 11
                        ProShep.Show
                        sheryl.Hide
                    Case 12
                        ProPit.Show
                        sheryl.Hide
                    Case 13
                        ProMtn.Show
                        sheryl.Hide
                    Case 14
                        Produch.Show
                        sheryl.Hide
                End Select
    End If
End Sub
