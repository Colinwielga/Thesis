VERSION 5.00
Begin VB.Form frmquiz 
   BackColor       =   &H80000012&
   Caption         =   "Form1"
   ClientHeight    =   7830
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   11100
   BeginProperty Font 
      Name            =   "Playbill"
      Size            =   15.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7830
   ScaleWidth      =   11100
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdgettoknow 
      Caption         =   "Get to Know the Characters"
      Height          =   495
      Left            =   8280
      TabIndex        =   30
      Top             =   1560
      Width           =   2655
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "Back to East High"
      BeginProperty Font 
         Name            =   "Playbill"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   8280
      TabIndex        =   29
      Top             =   120
      Width           =   2655
   End
   Begin VB.CommandButton cmdE4 
      Caption         =   "Breaking Free"
      Height          =   735
      Left            =   5520
      TabIndex        =   28
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton cmdE3 
      Caption         =   "We're all in This Together"
      Height          =   735
      Left            =   3840
      TabIndex        =   27
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton cmdE2 
      Caption         =   "Start of Something New"
      Height          =   735
      Left            =   2160
      TabIndex        =   26
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton cmdE1 
      Caption         =   "Bop to the Top"
      Height          =   735
      Left            =   360
      TabIndex        =   25
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CommandButton cmdD4 
      Caption         =   "Sports Pics"
      Height          =   735
      Left            =   5760
      TabIndex        =   24
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton cmdD3 
      Caption         =   "A Mirror"
      Height          =   735
      Left            =   3960
      TabIndex        =   23
      Top             =   3480
      Width           =   1695
   End
   Begin VB.CommandButton cmdD2 
      Caption         =   "Do Textbooks Count as Decoration?"
      Height          =   735
      Left            =   2040
      TabIndex        =   22
      Top             =   3480
      Width           =   1815
   End
   Begin VB.CommandButton cmdD1 
      Caption         =   "Sparkles, Sparkles, and more Sparkles"
      Height          =   735
      Left            =   360
      TabIndex        =   21
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton cmdC4 
      Caption         =   "Hangout with Family and Friends"
      Height          =   735
      Left            =   5520
      TabIndex        =   20
      Top             =   2400
      Width           =   1815
   End
   Begin VB.CommandButton cmdC3 
      Caption         =   "Keep Busy"
      Height          =   735
      Left            =   3960
      TabIndex        =   19
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton cmdC2 
      Caption         =   "Lose Yourself in a Book"
      Height          =   735
      Left            =   1920
      TabIndex        =   18
      Top             =   2400
      Width           =   1935
   End
   Begin VB.CommandButton cmdC1 
      Caption         =   "Get a Makeover"
      Height          =   735
      Left            =   360
      TabIndex        =   17
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton cmdB4 
      Caption         =   "Shooting Hoops"
      Height          =   495
      Left            =   6600
      TabIndex        =   16
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton cmdB3 
      Caption         =   "Following my Older Sibs Around"
      Height          =   495
      Left            =   3960
      TabIndex        =   15
      Top             =   1440
      Width           =   2535
   End
   Begin VB.CommandButton cmdB2 
      Caption         =   "Doing my Homework"
      Height          =   495
      Left            =   2040
      TabIndex        =   14
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton cmdB1 
      Caption         =   "Play Practice"
      Height          =   495
      Left            =   360
      TabIndex        =   13
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton cmdA4 
      Caption         =   "Gym"
      Height          =   495
      Left            =   4680
      TabIndex        =   12
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton cmdA3 
      Caption         =   "Lunch"
      Height          =   495
      Left            =   3240
      TabIndex        =   11
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton cmdA2 
      Caption         =   "Math"
      Height          =   495
      Left            =   1800
      TabIndex        =   10
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton cmdA1 
      BackColor       =   &H80000012&
      Caption         =   "Drama"
      Height          =   495
      Left            =   360
      TabIndex        =   9
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton cmdreset 
      Caption         =   "Reset Quiz"
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Who am I?"
      Height          =   1215
      Left            =   120
      TabIndex        =   7
      Top             =   5640
      Width           =   1215
   End
   Begin VB.PictureBox picresults2 
      Height          =   5415
      Left            =   8160
      ScaleHeight     =   5355
      ScaleWidth      =   2595
      TabIndex        =   6
      Top             =   2160
      Width           =   2655
   End
   Begin VB.PictureBox picresults 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   1440
      ScaleHeight     =   1875
      ScaleWidth      =   6435
      TabIndex        =   5
      Top             =   5640
      Width           =   6495
   End
   Begin VB.Label lbl5 
      BackColor       =   &H80000007&
      Caption         =   "5. What's your favorite HSM song?"
      BeginProperty Font 
         Name            =   "Playbill"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   4320
      Width           =   3135
   End
   Begin VB.Label lbl4 
      BackColor       =   &H80000007&
      Caption         =   "4. What's decorating your locker?"
      BeginProperty Font 
         Name            =   "Playbill"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   3120
      Width           =   3015
   End
   Begin VB.Label lbl3 
      BackColor       =   &H80000007&
      Caption         =   "3.What's the best way to get over a broken heart?"
      BeginProperty Font 
         Name            =   "Playbill"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   2040
      Width           =   4575
   End
   Begin VB.Label blb2 
      BackColor       =   &H80000007&
      Caption         =   "2. Where can you be found after school?"
      BeginProperty Font 
         Name            =   "Playbill"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   3735
   End
   Begin VB.Label lbl1 
      BackColor       =   &H80000007&
      Caption         =   "1. What's your favorite subject?"
      BeginProperty Font 
         Name            =   "Playbill"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "frmquiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: High School Musical
' Form name: Quiz
' Author: Laura Deal, Megan Haar, Kirsten Fasching
' Date Written: 10/28/08
'Objective: This program asks the user questions to find out which HSM character they are more like.
 'objective cont' there are 5 questions and they are compared with 4 of the HSM main characers
 ' objective cont' there is a tie breaker question (if needed) and a total is kept throughout the program

Option Explicit
Dim Total1 As Integer
Dim Total2 As Integer
Dim Total3 As Integer
Dim Total4 As Integer
Private Sub cmdA1_Click()
'sets the total for sharpay at zero
Total1 = 0

cmdA1.Visible = True
cmdA2.Visible = False
cmdA3.Visible = False
cmdA4.Visible = False
    
'keeps a total for questions answered similar to Sharpay
Total1 = Total1 + 1
End Sub

Private Sub cmdA2_Click()
'sets the total for Gabriella at zero
Total2 = 0
cmdA2.Visible = True
cmdA3.Visible = False
cmdA4.Visible = False
cmdA1.Visible = False

'keeps a total for questions answered similar to gabriella
Total2 = Total2 + 1
End Sub

Private Sub cmdA3_Click()
'sets the total for Ryan at zero
Total3 = 0
cmdA3.Visible = True
cmdA4.Visible = False
cmdA1.Visible = False
cmdA2.Visible = False

'keeps a total for questions answered similar to Ryan
Total3 = Total3 + 1
End Sub

Private Sub cmdA4_Click()
'sets the total for Troy at zero
Total4 = 0
cmdA4.Visible = True
cmdA1.Visible = False
cmdA2.Visible = False
cmdA3.Visible = False

'keeps a total for questions answered similar to Troy
Total4 = Total4 + 1
End Sub

Private Sub cmdB1_Click()
cmdB1.Visible = True
cmdB2.Visible = False
cmdB3.Visible = False
cmdB4.Visible = False
    
'keeps a total for questions answered similar to Sharpay
Total1 = Total1 + 1
End Sub

Private Sub cmdB2_Click()
cmdB2.Visible = True
cmdB3.Visible = False
cmdB4.Visible = False
cmdB1.Visible = False

'keeps a total for questions answered similar to Gabriella
Total2 = Total2 + 1
End Sub

Private Sub cmdB3_Click()
cmdB3.Visible = True
cmdB4.Visible = False
cmdB1.Visible = False
cmdB2.Visible = False

'keeps a total for questions answered similar to Ryan
Total3 = Total3 + 1
End Sub

Private Sub cmdB4_Click()
cmdB4.Visible = True
cmdB1.Visible = False
cmdB2.Visible = False
cmdB3.Visible = False

'keeps a total for questions answered similar to Troy
Total4 = Total4 + 1
End Sub

Private Sub cmdback_Click()
'brings the user back to the buttons page to go to another activity or leave
frmauthors.Hide
frmbuttons.Show
Frmcharacter.Hide
frmtitle.Hide
frmnamethattune.Hide
FrmTrivia.Hide
frmquiz.Hide
End Sub

Private Sub cmdC1_Click()
cmdC1.Visible = True
cmdC2.Visible = False
cmdC3.Visible = False
cmdC4.Visible = False
    
'keeps a total for questions answered similar to Sharpay
Total1 = Total1 + 1
End Sub

Private Sub cmdC2_Click()
cmdC2.Visible = True
cmdC3.Visible = False
cmdC4.Visible = False
cmdC1.Visible = False

'keeps a total for questions answered similar to Gabriella
Total2 = Total2 + 1
End Sub

Private Sub cmdC3_Click()
cmdC3.Visible = True
cmdC4.Visible = False
cmdC1.Visible = False
cmdC2.Visible = False

'keeps a total for questions answered similar to Ryan
Total3 = Total3 + 1
End Sub

Private Sub cmdC4_Click()
cmdC4.Visible = True
cmdC1.Visible = False
cmdC2.Visible = False
cmdC3.Visible = False

'keeps a total for questions answered similar to Troy
Total4 = Total4 + 1
End Sub

Private Sub cmdD1_Click()
cmdD1.Visible = True
cmdD2.Visible = False
cmdD3.Visible = False
cmdD4.Visible = False
    
'keeps a total for questions answered similar to Sharpay
Total1 = Total1 + 1
End Sub

Private Sub cmdD2_Click()
cmdD2.Visible = True
cmdD3.Visible = False
cmdD4.Visible = False
cmdD1.Visible = False

'keeps a total for questions answered similar to Gabriella
Total2 = Total2 + 1
End Sub

Private Sub cmdD3_Click()
cmdD3.Visible = True
cmdD4.Visible = False
cmdD1.Visible = False
cmdD2.Visible = False

'keeps a total for questions answered similar to Ryan
Total3 = Total3 + 1
End Sub

Private Sub cmdD4_Click()
cmdD4.Visible = True
cmdD1.Visible = False
cmdD2.Visible = False
cmdD3.Visible = False

'keeps a total for questions answered similar to Troy
Total4 = Total4 + 1
End Sub

Private Sub cmdE1_Click()
cmdE1.Visible = True
cmdE2.Visible = False
cmdE3.Visible = False
cmdE4.Visible = False
    
'keeps a total for questions answered similar to Sharpay
Total1 = Total1 + 1
End Sub

Private Sub cmdE2_Click()
cmdE2.Visible = True
cmdE3.Visible = False
cmdE4.Visible = False
cmdE1.Visible = False

'keeps a total for questions answered similar to Gabriella
Total2 = Total2 + 1
End Sub

Private Sub cmdE3_Click()
cmdE3.Visible = True
cmdE4.Visible = False
cmdE1.Visible = False
cmdE2.Visible = False

'keeps a total for questions answered similar to Ryan
Total3 = Total3 + 1
End Sub

Private Sub cmdE4_Click()
cmdE4.Visible = True
cmdE1.Visible = False
cmdE2.Visible = False
cmdE3.Visible = False

'keeps a total for questions answered similar to Troy
Total4 = Total4 + 1
End Sub

Private Sub cmdgettoknow_Click()
'Brings the user to the 'Get to Know the Characters" page to find out more about the person they are most similar to
frmauthors.Hide
frmbuttons.Hide
Frmcharacter.Show
frmtitle.Hide
frmnamethattune.Hide
FrmTrivia.Hide
frmquiz.Hide
frmtitle.Hide
End Sub

Private Sub cmdreset_Click()
'clear pic boxes
picresults.Cls

'clear counters
Total1 = 0
Total2 = 0
Total3 = 0
Total4 = 0

'make buttons visible
cmdA1.Visible = True
cmdA2.Visible = True
cmdA3.Visible = True
cmdA4.Visible = True

cmdB4.Visible = True
cmdB1.Visible = True
cmdB2.Visible = True
cmdB3.Visible = True

cmdC2.Visible = True
cmdC3.Visible = True
cmdC4.Visible = True
cmdC1.Visible = True

cmdD1.Visible = True
cmdD2.Visible = True
cmdD3.Visible = True
cmdD4.Visible = True

cmdE4.Visible = True
cmdE1.Visible = True
cmdE2.Visible = True
cmdE3.Visible = True
End Sub

Private Sub Command1_Click()
Dim Tie As String
Dim highest As Integer
Dim total As String

picresults.Cls

'Results

If Total1 > Total2 And Total1 > Total3 And Total1 > Total4 Then
    highest = Total1
    picresults.Print "You're Sharpay!!"
    picresults.Print "Go to the Get to Know the Characters page to learn more!"
    picresults2.Picture = LoadPicture(App.Path & "\Sharpay(quiz).jpg")
End If

If Total2 > Total1 And Total2 > Total3 And Total2 > Total4 Then
    highest = Total2
    picresults.Print "You're Gabriella!!"
    picresults.Print "Go to the Get to Know the Characters page to learn more!"
    picresults2.Picture = LoadPicture(App.Path & "\Gabriella(quiz).jpg")
End If

If Total3 > Total1 And Total3 > Total2 And Total3 > Total4 Then
    highest = Total3
    picresults.Print "You're Ryan!!"
    picresults.Print "Go to the Get to Know the Characters page to learn more!"
    picresults2.Picture = LoadPicture(App.Path & "\Ryan(Quiz).jpg")
End If

If Total4 > Total1 And Total4 > Total2 And Total4 > Total3 Then
    highest = Total4
    picresults.Print "You're Troy!!"
    picresults.Print "Go to the Get to Know the Characters page to learn more!"
    picresults2.Picture = LoadPicture(App.Path & "\Troy(Quiz).jpg")
End If

'tie breakers

If highest = 0 Then
    If (Total1 = Total2) Then
        Tie = InputBox("Pink or Green?", "Tie-Breaker!")
            If LCase(Tie) = "pink" Then
                picresults.Print "You're Sharpay!!"
                picresults.Print "Go to the Get to Know the Characters page to learn more!"
                picresults2.Picture = LoadPicture(App.Path & "\Sharpay(quiz).jpg")
            Else
                picresults.Print "You're Gabriella!!"
                picresults.Print "Go to the Get to Know the Characters page to learn more!"
                picresults2.Picture = LoadPicture(App.Path & "\Gabriella(quiz).jpg")
            End If
    ElseIf (Total1 = Total3) Then
        Tie = InputBox("Hat or Headband?", "Tie-Breaker!")
            If LCase(Tie) = "headband" Then
                picresults.Print "You're Sharpay!!"
                picresults.Print "Go to the Get to Know the Characters page to learn more!"
                picresults2.Picture = LoadPicture(App.Path & "\Sharpay(quiz).jpg")
            Else
                picresults.Print "You're Ryan!!"
                picresults.Print "Go to the Get to Know the Characters page to learn more!"
                picresults2.Picture = LoadPicture(App.Path & "\Ryan(Quiz).jpg")
            End If
    ElseIf (Total1 = Total4) Then
        Tie = InputBox("Basketball or Acting?", "Tie-Breaker!")
            If LCase(Tie) = "basketball" Then
                picresults.Print "You're Troy!!"
                picresults.Print "Go to the Get to Know the Characters page to learn more!"
                picresults2.Picture = LoadPicture(App.Path & "\Troy(Quiz).jpg")
            Else
                picresults.Print "You're Sharpay!!"
                picresults.Print "Go to the Get to Know the Characters page to learn more!"
                picresults2.Picture = LoadPicture(App.Path & "\Sharpay(quiz).jpg")
            End If
    ElseIf (Total2 = Total3) Then
        Tie = InputBox("Red or Orange?", "Tie-Breaker!")
            If LCase(Tie) = "red" Then
                picresults.Print "You're Gabriella!!"
                picresults.Print "Go to the Get to Know the Characters page to learn more!"
                picresults2.Picture = LoadPicture(App.Path & "\Gabriella(quiz).jpg")
            Else
                picresults.Print "You're Ryan!!"
                picresults.Print "Go to the Get to Know the Characters page to learn more!"
                picresults2.Picture = LoadPicture(App.Path & "\Ryan(Quiz).jpg")
            End If
    ElseIf (Total2 = Total4) Then
        Tie = InputBox("Homework or Baketball?", "Tie-Breaker!")
            If LCase(Tie) = "homework" Then
                picresults.Print "You're Gabriella!!"
                picresults.Print "Go to the Get to Know the Characters page to learn more!"
                picresults2.Picture = LoadPicture(App.Path & "\Gabriella(quiz).jpg")
            Else
                picresults.Print "You're Troy!!"
                picresults.Print "Go to the Get to Know the Characters page to learn more!"
                picresults2.Picture = LoadPicture(App.Path & "\Troy(Quiz).jpg")
            End If
    ElseIf (Total3 = Total4) Then
        Tie = InputBox("Dancing or Basketball?", "Tie-Breaker!")
            If LCase(Tie) = "basketball" Then
                picresults.Print "You're Troy!!"
                picresults.Print "Go to the Get to Know the Characters page to learn more!"
                picresults2.Picture = LoadPicture(App.Path & "\Troy(Quiz).jpg")
            Else
                picresults.Print "You're Ryan!!"
                picresults.Print "Go to the Get to Know the Characters page to learn more!"
                picresults2.Picture = LoadPicture(App.Path & "\Ryan(Quiz).jpg")
            End If
    End If
End If

End Sub
