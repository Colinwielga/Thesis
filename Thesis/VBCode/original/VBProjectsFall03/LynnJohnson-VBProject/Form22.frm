VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H008080FF&
   Caption         =   "Form1"
   ClientHeight    =   10125
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11985
   LinkTopic       =   "Form1"
   ScaleHeight     =   10125
   ScaleWidth      =   11985
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdreturn 
      Caption         =   "Return to Menu"
      Height          =   1095
      Left            =   9960
      TabIndex        =   37
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdsixteen 
      Caption         =   "16"
      Height          =   1815
      Left            =   6960
      TabIndex        =   35
      Top             =   7800
      Width           =   1935
   End
   Begin VB.PictureBox picresults16 
      Height          =   1815
      Left            =   6960
      Picture         =   "Form22.frx":0000
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   28
      Top             =   7800
      Width           =   1935
   End
   Begin VB.CommandButton cmdfifteen 
      Caption         =   "15"
      Height          =   1815
      Left            =   4680
      TabIndex        =   34
      Top             =   7800
      Width           =   1935
   End
   Begin VB.CommandButton cmdfourteen 
      Caption         =   "14"
      Height          =   1815
      Left            =   2520
      TabIndex        =   33
      Top             =   7800
      Width           =   1815
   End
   Begin VB.CommandButton cmdthirteen 
      Caption         =   "13"
      Height          =   1815
      Left            =   360
      TabIndex        =   32
      Top             =   7800
      Width           =   1815
   End
   Begin VB.CommandButton cmdtwelve 
      Caption         =   "12"
      Height          =   1815
      Left            =   6960
      TabIndex        =   31
      Top             =   5640
      Width           =   1935
   End
   Begin VB.CommandButton cmdeleven 
      Caption         =   "11"
      Height          =   1815
      Left            =   4680
      TabIndex        =   30
      Top             =   5640
      Width           =   1935
   End
   Begin VB.CommandButton cmdten 
      Caption         =   "10"
      Height          =   1815
      Left            =   2520
      TabIndex        =   29
      Top             =   5640
      Width           =   1815
   End
   Begin VB.PictureBox picresults14 
      Height          =   1815
      Left            =   2520
      Picture         =   "Form22.frx":4C84
      ScaleHeight     =   1755
      ScaleWidth      =   1755
      TabIndex        =   26
      Top             =   7800
      Width           =   1815
   End
   Begin VB.PictureBox picresults13 
      Height          =   1815
      Left            =   360
      Picture         =   "Form22.frx":9912
      ScaleHeight     =   1755
      ScaleWidth      =   1755
      TabIndex        =   25
      Top             =   7800
      Width           =   1815
   End
   Begin VB.PictureBox picresults12 
      Height          =   1815
      Left            =   6960
      Picture         =   "Form22.frx":E5A0
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   24
      Top             =   5640
      Width           =   1935
   End
   Begin VB.PictureBox picresults11 
      Height          =   1815
      Left            =   4680
      Picture         =   "Form22.frx":12D1F
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   23
      Top             =   5640
      Width           =   1935
   End
   Begin VB.PictureBox picresults10 
      Height          =   1815
      Left            =   2520
      Picture         =   "Form22.frx":1749E
      ScaleHeight     =   1755
      ScaleWidth      =   1755
      TabIndex        =   22
      Top             =   5640
      Width           =   1815
   End
   Begin VB.CommandButton cmdnine 
      Caption         =   "9"
      Height          =   1815
      Left            =   360
      TabIndex        =   21
      Top             =   5640
      Width           =   1815
   End
   Begin VB.CommandButton cmdeight 
      Caption         =   "8"
      Height          =   1815
      Left            =   6960
      TabIndex        =   20
      Top             =   3480
      Width           =   1935
   End
   Begin VB.CommandButton cmdseven 
      Caption         =   "7"
      Height          =   1815
      Left            =   4680
      TabIndex        =   19
      Top             =   3480
      Width           =   1935
   End
   Begin VB.CommandButton cmdsix 
      Caption         =   "6"
      Height          =   1815
      Left            =   2520
      TabIndex        =   18
      Top             =   3480
      Width           =   1815
   End
   Begin VB.PictureBox picresults9 
      Height          =   1815
      Left            =   360
      Picture         =   "Form22.frx":1DEA7
      ScaleHeight     =   1755
      ScaleWidth      =   1755
      TabIndex        =   17
      Top             =   5640
      Width           =   1815
   End
   Begin VB.PictureBox picresults8 
      Height          =   1815
      Left            =   6960
      Picture         =   "Form22.frx":248B0
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   16
      Top             =   3480
      Width           =   1935
   End
   Begin VB.PictureBox picresults7 
      Height          =   1815
      Left            =   4680
      Picture         =   "Form22.frx":2A2A5
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   15
      Top             =   3480
      Width           =   1935
   End
   Begin VB.PictureBox picresults6 
      Height          =   1815
      Left            =   2520
      Picture         =   "Form22.frx":2FC9A
      ScaleHeight     =   1755
      ScaleWidth      =   1755
      TabIndex        =   14
      Top             =   3480
      Width           =   1815
   End
   Begin VB.CommandButton cmdfive 
      Caption         =   "Memory5"
      Height          =   1815
      Left            =   360
      TabIndex        =   13
      Top             =   3480
      Width           =   1815
   End
   Begin VB.PictureBox picresults5 
      Height          =   1815
      Left            =   360
      Picture         =   "Form22.frx":34EC1
      ScaleHeight     =   1755
      ScaleWidth      =   1755
      TabIndex        =   12
      Top             =   3480
      Width           =   1815
   End
   Begin VB.CommandButton cmdplay 
      Caption         =   "Play again"
      Height          =   975
      Left            =   9960
      TabIndex        =   11
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmdfour 
      Caption         =   "Memory4"
      Height          =   1815
      Left            =   6960
      TabIndex        =   10
      Top             =   1320
      Width           =   1935
   End
   Begin VB.CommandButton cmdthree 
      Caption         =   "Memory3"
      Height          =   1815
      Left            =   4680
      Picture         =   "Form22.frx":3A0E8
      TabIndex        =   9
      Top             =   1320
      Width           =   1935
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "Clear Match"
      Height          =   975
      Left            =   9960
      TabIndex        =   8
      Top             =   3480
      Width           =   1215
   End
   Begin VB.PictureBox pbxresults 
      BackColor       =   &H00C0E0FF&
      Height          =   855
      Left            =   5760
      ScaleHeight     =   795
      ScaleWidth      =   5115
      TabIndex        =   7
      Top             =   120
      Width           =   5175
   End
   Begin VB.PictureBox picresults4 
      Height          =   1815
      Left            =   6960
      Picture         =   "Form22.frx":3BCB5
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   6
      Top             =   1320
      Width           =   1935
   End
   Begin VB.PictureBox picresults3 
      Height          =   1815
      Left            =   4680
      Picture         =   "Form22.frx":4146F
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   5
      Top             =   1320
      Width           =   1935
   End
   Begin VB.CommandButton cmdtwo 
      Caption         =   "Memory2"
      Height          =   1815
      Left            =   2520
      TabIndex        =   4
      Top             =   1320
      Width           =   1815
   End
   Begin VB.PictureBox picresults2 
      Height          =   1815
      Left            =   2520
      Picture         =   "Form22.frx":46C29
      ScaleHeight     =   1755
      ScaleWidth      =   1755
      TabIndex        =   3
      Top             =   1320
      Width           =   1815
   End
   Begin VB.CommandButton cmdone 
      BackColor       =   &H00FF8080&
      Caption         =   "Memory1"
      Height          =   1815
      Left            =   360
      TabIndex        =   2
      Top             =   1320
      Width           =   1815
   End
   Begin VB.PictureBox picresults1 
      Height          =   1815
      Left            =   360
      Picture         =   "Form22.frx":4B988
      ScaleHeight     =   1755
      ScaleWidth      =   1755
      TabIndex        =   1
      Top             =   1320
      Width           =   1815
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   975
      Left            =   9960
      TabIndex        =   0
      Top             =   6360
      Width           =   1215
   End
   Begin VB.PictureBox picresults15 
      Height          =   1815
      Left            =   4680
      Picture         =   "Form22.frx":506E7
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   27
      Top             =   7800
      Width           =   1935
   End
   Begin VB.Label lblinstruction 
      BackColor       =   &H00C0E0FF&
      Caption         =   $"Form22.frx":5536B
      ForeColor       =   &H00404080&
      Height          =   615
      Left            =   360
      TabIndex        =   36
      Top             =   240
      Width           =   4575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim icount As Integer


Private Sub cmdclear_Click()
    If cmdtwo.Visible = False And cmdone.Visible = False Then
        picresults1.Visible = False
        picresults2.Visible = False
    ElseIf cmdthree.Visible = False And cmdfour.Visible = False Then
        picresults3.Visible = False
        picresults4.Visible = False
    ElseIf cmdfive.Visible = False And cmdsix.Visible = False Then
        picresults5.Visible = False
        picresults6.Visible = False
    ElseIf cmdseven.Visible = False And cmdeight.Visible = False Then
        picresults7.Visible = False
        picresults8.Visible = False
    ElseIf cmdnine.Visible = False And cmdten.Visible = False Then
        picresults9.Visible = False
        picresults10.Visible = False
    ElseIf cmdeleven.Visible = False And cmdtwelve.Visible = False Then
        picresults11.Visible = False
        picresults12.Visible = False
    ElseIf cmdthirteen.Visible = False And cmdfourteen.Visible = False Then
        picresults13.Visible = False
        picresults14.Visible = False
    ElseIf cmdfifteen.Visible = False And cmdsixteen.Visible = False Then
        picresults15.Visible = False
        picresults16.Visible = False
    Else
        pbxresults.Cls
        pbxresults.Print "This is not a match.  Click on the cards to cover them again and pick again."
        
    End If
    
End Sub

Private Sub cmdeight_Click()
    icount = icount + 1
    pbxresults.Cls
    cmdeight.Visible = False
    picresults8.Visible = True
    If cmdseven.Visible = False Then
        pbxresults.Print "You found another dog! You found a match!"
    Else
        pbxresults.Print "You found a dog!"
    End If
End Sub

Private Sub cmdeleven_Click()
    icount = icount + 1
    pbxresults.Cls
    cmdeleven.Visible = False
    picresults11.Visible = True
    If cmdtwelve.Visible = False Then
        pbxresults.Print "You found another cow! You found a match!"
    Else
        pbxresults.Print "You found a cow!"
    End If
End Sub

Private Sub cmdfifteen_Click()
    icount = icount + 1
    pbxresults.Cls
    cmdfifteen.Visible = False
    picresults15.Visible = True
    If cmdsixteen.Visible = False Then
        pbxresults.Print "You found another chick! You found a match!"
    Else
        pbxresults.Print "You found a baby chick!"
    End If
End Sub

Private Sub cmdfive_Click()
    icount = icount + 1
    pbxresults.Cls
    cmdfive.Visible = False
    picresults5.Visible = True
    If cmdsix.Visible = False Then
        pbxresults.Print "You found the other bird! You found a match!"
    Else
        pbxresults.Print "You found a bird!"
    End If
End Sub

Private Sub cmdfour_Click()
    icount = icount + 1
    pbxresults.Cls
    cmdfour.Visible = False
    picresults4.Visible = True
    If cmdthree.Visible = False Then
        pbxresults.Print "You found the other elephant! You found a match!"
    Else
        pbxresults.Print "You found an elephant!"
    End If
End Sub

Private Sub cmdfourteen_Click()
    icount = icount + 1
    pbxresults.Cls
    cmdfourteen.Visible = False
    picresults14.Visible = True
    If cmdthirteen.Visible = False Then
        pbxresults.Print "You found another fish! You found a match!"
    Else
        pbxresults.Print "You found a fish!"
    End If
End Sub

Private Sub cmdnine_Click()
    icount = icount + 1
    pbxresults.Cls
    cmdnine.Visible = False
    picresults9.Visible = True
    If cmdten.Visible = False Then
        pbxresults.Print "You found another butterfly! You found a match!"
    Else
        pbxresults.Print "You found a butterfly!"
    End If
End Sub

Private Sub cmdone_Click()
    icount = icount + 1
    pbxresults.Cls
    cmdone.Visible = False
    picresults1.Visible = True
    If cmdtwo.Visible = False Then
        pbxresults.Print "You found the other frog! You found a match!"
    Else
        pbxresults.Print "You found a frog!"
    End If
    
End Sub

Private Sub cmdplay_Click()
    Form2.Show
    Form1.Hide
    
End Sub

Private Sub cmdquit_Click()
    End
    
End Sub

Private Sub cmdreturn_Click()
    Form3.Show
    Form1.Hide
    
End Sub

Private Sub cmdseven_Click()
    icount = icount + 1
    pbxresults.Cls
    cmdseven.Visible = False
    picresults7.Visible = True
    If cmdeight.Visible = False Then
        pbxresults.Print "You found another dog! You found a match!"
    Else
        pbxresults.Print "You found a dog!"
    End If
End Sub

Private Sub cmdsix_Click()
    icount = icount + 1
    pbxresults.Cls
    cmdsix.Visible = False
    picresults6.Visible = True
    If cmdfive.Visible = False Then
        pbxresults.Print "You found the other bird! You found a match!"
    Else
        pbxresults.Print "You found a bird!"
    End If
End Sub

Private Sub cmdsixteen_Click()
    icount = icount + 1
    pbxresults.Cls
    cmdsixteen.Visible = False
    picresults16.Visible = True
    If cmdfifteen.Visible = False Then
        pbxresults.Print "You found another chick! You found a match!"
    Else
        pbxresults.Print "You found a baby chick!"
    End If
End Sub

Private Sub cmdten_Click()
    icount = icount + 1
    pbxresults.Cls
    cmdten.Visible = False
    picresults10.Visible = True
    If cmdnine.Visible = False Then
        pbxresults.Print "You found another butterfly! You found a match!"
    Else
        pbxresults.Print "You found a butterfly!"
    End If
End Sub

Private Sub cmdthirteen_Click()
    icount = icount + 1
    pbxresults.Cls
    cmdthirteen.Visible = False
    picresults13.Visible = True
    If cmdfourteen.Visible = False Then
        pbxresults.Print "You found another fish! You found a match!"
    Else
        pbxresults.Print "You found a fish!"
    End If
End Sub

Private Sub cmdthree_Click()
    icount = icount + 1
    pbxresults.Cls
    cmdthree.Visible = False
    picresults3.Visible = True
    If cmdfour.Visible = False Then
        pbxresults.Print "You found the other elephant! You found a match!"
    Else
        pbxresults.Print "You found an elephant!"
    End If
    
End Sub

Private Sub cmdtwelve_Click()
    icount = icount + 1
    pbxresults.Cls
    cmdtwelve.Visible = False
    picresults12.Visible = True
    If cmdeleven.Visible = False Then
        pbxresults.Print "You found another cow! You found a match!"
    Else
        pbxresults.Print "You found a cow!"
    End If
End Sub

Private Sub cmdtwo_Click()
    icount = icount + 1
    pbxresults.Cls
    cmdtwo.Visible = False
    picresults2.Visible = True
    If cmdone.Visible = False Then
        pbxresults.Print "You found the other frog!  You found a match!"
    Else
        pbxresults.Print "You found a frog!"
    End If
    
End Sub

Private Sub picresults10_Click()
    cmdten.Visible = True
    picresults10.Visible = False
    
End Sub

Private Sub picresults1_Click()
    cmdone.Visible = True
    picresults1.Visible = False
    
End Sub


Private Sub picresults11_Click()
    cmdeleven.Visible = True
    picresults11.Visible = False
    
End Sub

Private Sub picresults12_Click()
    cmdtwelve.Visible = True
    picresults12.Visible = False
    
End Sub

Private Sub picresults14_Click()
    cmdfourteen.Visible = True
    picresults14.Visible = False
    
End Sub

Private Sub picresults15_Click()
    cmdfifteen.Visible = True
    picresults15.Visible = False
    
End Sub

Private Sub picresults16_Click()
    cmdsixteen.Visible = True
    picresults16.Visible = False
    
End Sub

Private Sub picresults2_Click()
    cmdtwo.Visible = True
    picresults2.Visible = False
    
End Sub

Private Sub picresults3_Click()
    cmdthree.Visible = True
    picresults3.Visible = False
    
End Sub

Private Sub picresults4_Click()
    cmdfour.Visible = True
    picresults4.Visible = False
    
End Sub

Private Sub picresults5_Click()
    cmdfive.Visible = True
    picresults5.Visible = False
    
End Sub

Private Sub picresults6_Click()
    cmdsix.Visible = True
    picresults6.Visible = False
    
End Sub

Private Sub picresults7_Click()
    cmdseven.Visible = True
    picresults7.Visible = False
    
End Sub

Private Sub picresults8_Click()
    cmdeight.Visible = True
    picresults8.Visible = False
    
End Sub

Private Sub picresults9_Click()
    cmdnine.Visible = True
    picresults9.Visible = False
    
End Sub

Private Sub picresults13_Click()
    cmdthirteen.Visible = True
    picresults13.Visible = False
    
End Sub
