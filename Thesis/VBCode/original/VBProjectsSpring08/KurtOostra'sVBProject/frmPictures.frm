VERSION 5.00
Begin VB.Form frmPictures 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox lblInfo 
      Height          =   375
      Left            =   4800
      ScaleHeight     =   315
      ScaleWidth      =   5235
      TabIndex        =   3
      Top             =   9360
      Width           =   5295
   End
   Begin VB.CommandButton cmdSlideShow 
      BackColor       =   &H000000FF&
      Caption         =   "Slide show"
      Height          =   1575
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   360
      Width           =   2415
   End
   Begin VB.PictureBox picResults 
      Height          =   9015
      Left            =   2760
      ScaleHeight     =   8955
      ScaleWidth      =   8940
      TabIndex        =   1
      Top             =   240
      Width           =   9000
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Return to Main Menu"
      Height          =   1215
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6960
      Width           =   2055
   End
End
Attribute VB_Name = "frmpictures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Theater Lighting
'Form Name: frmPictures
'Author: Kurt Oostra
'Date Written:3/27/08
'Objective: View a slideshow of different shows
Option Explicit
Private Sub cmdReturn_Click()
'Returns to main Menu
frmMainMenu.Show
frmpictures.Hide
End Sub

Private Sub cmdSlideShow_Click()
Dim j As Integer, ctr As Integer, picnames(1 To 7) As String, info(1 To 7) As String
Dim pick As Integer, stopper As Integer, t As Double, old As Integer
'Loads pictures into an array called picnames
Open App.Path & "\picNames.txt" For Input As #1
    ctr = 0
    Do While Not EOF(1)
        ctr = ctr + 1
        Input #1, picnames(ctr), info(ctr)
Loop
Close #1
pick = 1
stopper = 0
lblInfo.Visible = True
'starts the slideshow
Do While (stopper < 6)
    lblInfo.Cls
    lblInfo.Print info(pick)
    picResults.Picture = LoadPicture(App.Path & "\" & picnames(pick))
    t = Timer
    'times the lenght the picture is up
    Do While (Timer - t) < 2
    Loop
    stopper = stopper + 1
    old = pick
    pick = (stopper Mod ctr) + 1
Loop
frmpictures.Show
End Sub
