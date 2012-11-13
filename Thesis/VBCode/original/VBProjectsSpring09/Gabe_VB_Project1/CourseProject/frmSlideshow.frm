VERSION 5.00
Begin VB.Form frmSlideshow 
   BackColor       =   &H00FF0000&
   Caption         =   "Slideshow"
   ClientHeight    =   4680
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7080
   LinkTopic       =   "Form1"
   Picture         =   "frmSlideshow.frx":0000
   ScaleHeight     =   4680
   ScaleWidth      =   7080
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picResults 
      BackColor       =   &H000000C0&
      Height          =   7455
      Left            =   6000
      ScaleHeight     =   7395
      ScaleWidth      =   7995
      TabIndex        =   2
      Top             =   1560
      Width           =   8055
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Start the Slideshow!"
      Height          =   735
      Left            =   5520
      TabIndex        =   1
      Top             =   720
      Width           =   2775
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Back to Main Menu"
      Height          =   855
      Left            =   8640
      Picture         =   "frmSlideshow.frx":DF41
      TabIndex        =   0
      Top             =   600
      Width           =   1815
   End
End
Attribute VB_Name = "frmSlideshow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Fun with CSB/SJU History!
'frmSlideshow
'Audrey Gabe
'Written 3/24/09
'This section does a slideshow of CSB/SJU archival pictures

Private Sub cmdMenu_Click()
frmSlideshow.Hide
frmMenu.Show
End Sub

Private Sub cmdPlay_Click()
Dim stopper As Integer, t As Double, ctr2 As Double, PicCTR As Integer, pictures(1 To 100) As String, Counter As Integer
Counter = 0
Open App.Path & "\slideshow.txt" For Input As #1 'Opens picture file
    Do While Not EOF(1) 'Reads file
        Counter = Counter + 1
        Input #1, pictures(Counter)
    Loop
PicCTR = 1

stopper = 0
Do While (stopper < 19) 'Displays each picture until the stopper reaches 19
    picResults.Picture = LoadPicture(App.Path & "\" & pictures(PicCTR)) 'displays new picture
    
    t = Timer
    Do While (Timer - t) < 2
        ctr2 = ctr2 + 1
        If ctr2 = 1000000 Then
            ctr2 = 0
        End If
    Loop
    
    stopper = stopper + 1
    
    PicCTR = (stopper Mod Counter) + 1
Loop
picResults.Cls
Close #1
   
frmSlideshow.Hide
frmMenu.Show
End Sub
