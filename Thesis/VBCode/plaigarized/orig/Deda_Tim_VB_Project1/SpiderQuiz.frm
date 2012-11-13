VERSION 5.00
Begin VB.Form Spider_Quiz 
   Caption         =   "Form1"
   ClientHeight    =   11445
   ClientLeft      =   3090
   ClientTop       =   2820
   ClientWidth     =   13980
   LinkTopic       =   "Form1"
   ScaleHeight     =   11445
   ScaleWidth      =   13980
   Begin VB.PictureBox PicResult8 
      BackColor       =   &H80000001&
      ForeColor       =   &H000000FF&
      Height          =   1935
      Left            =   6240
      ScaleHeight     =   1875
      ScaleWidth      =   1995
      TabIndex        =   31
      Top             =   7920
      Width           =   2055
   End
   Begin VB.PictureBox PicResult1 
      BackColor       =   &H80000001&
      ForeColor       =   &H000000FF&
      Height          =   1935
      Left            =   840
      ScaleHeight     =   1875
      ScaleWidth      =   1995
      TabIndex        =   30
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton Show1 
      BackColor       =   &H000000FF&
      Caption         =   "Start"
      Height          =   615
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   10200
      Width           =   2055
   End
   Begin VB.CommandButton Instructions 
      BackColor       =   &H000000FF&
      Caption         =   "Instructions"
      Height          =   615
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   10200
      Width           =   1695
   End
   Begin VB.PictureBox PicResult6 
      BackColor       =   &H80000001&
      ForeColor       =   &H000000FF&
      Height          =   1935
      Left            =   6240
      ScaleHeight     =   1875
      ScaleWidth      =   1995
      TabIndex        =   18
      Top             =   5640
      Width           =   2055
   End
   Begin VB.PictureBox PicResult4 
      BackColor       =   &H80000001&
      ForeColor       =   &H000000FF&
      Height          =   1935
      Left            =   6240
      ScaleHeight     =   1875
      ScaleWidth      =   1995
      TabIndex        =   17
      Top             =   3360
      Width           =   2055
   End
   Begin VB.PictureBox PicResult2 
      BackColor       =   &H80000001&
      ForeColor       =   &H000000FF&
      Height          =   1935
      Left            =   6240
      ScaleHeight     =   1875
      ScaleWidth      =   1995
      TabIndex        =   16
      Top             =   1080
      Width           =   2055
   End
   Begin VB.PictureBox PicResult7 
      BackColor       =   &H80000001&
      ForeColor       =   &H000000FF&
      Height          =   1935
      Left            =   840
      ScaleHeight     =   1875
      ScaleWidth      =   1995
      TabIndex        =   15
      Top             =   7920
      Width           =   2055
   End
   Begin VB.PictureBox PicResult5 
      BackColor       =   &H80000001&
      ForeColor       =   &H000000FF&
      Height          =   1935
      Left            =   840
      ScaleHeight     =   1875
      ScaleWidth      =   1995
      TabIndex        =   14
      Top             =   5640
      Width           =   2055
   End
   Begin VB.PictureBox PicResult3 
      BackColor       =   &H80000001&
      ForeColor       =   &H000000FF&
      Height          =   1935
      Left            =   840
      ScaleHeight     =   1875
      ScaleWidth      =   1995
      TabIndex        =   13
      Top             =   3360
      Width           =   2055
   End
   Begin VB.CommandButton Results 
      BackColor       =   &H000000FF&
      Caption         =   "Show Results"
      Height          =   615
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   10200
      Width           =   2055
   End
   Begin VB.PictureBox PicResults 
      BackColor       =   &H80000001&
      ForeColor       =   &H000000FF&
      Height          =   9015
      Left            =   11160
      ScaleHeight     =   8955
      ScaleWidth      =   2355
      TabIndex        =   10
      Top             =   1440
      Width           =   2415
   End
   Begin VB.TextBox Ans3 
      Height          =   375
      Left            =   3120
      TabIndex        =   9
      Top             =   3360
      Width           =   2175
   End
   Begin VB.TextBox Ans4 
      Height          =   375
      Left            =   8520
      TabIndex        =   8
      Top             =   3360
      Width           =   2175
   End
   Begin VB.TextBox Ans5 
      Height          =   375
      Left            =   3120
      TabIndex        =   7
      Top             =   5640
      Width           =   2175
   End
   Begin VB.TextBox Ans6 
      Height          =   375
      Left            =   8520
      TabIndex        =   6
      Top             =   5640
      Width           =   2175
   End
   Begin VB.TextBox Ans7 
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   7920
      Width           =   2175
   End
   Begin VB.TextBox Ans8 
      Height          =   375
      Left            =   8520
      TabIndex        =   4
      Top             =   7920
      Width           =   2175
   End
   Begin VB.TextBox Ans2 
      Height          =   375
      Left            =   8520
      TabIndex        =   3
      Top             =   1080
      Width           =   2175
   End
   Begin VB.TextBox Ans1 
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   1080
      Width           =   2175
   End
   Begin VB.CommandButton MainReturn 
      BackColor       =   &H000000FF&
      Caption         =   "Return to Main Menu"
      Height          =   735
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   10200
      Width           =   1935
   End
   Begin VB.CommandButton SpiderReturn 
      BackColor       =   &H000000FF&
      Caption         =   "Return to Spiderman"
      Height          =   735
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   10200
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Name That Character From the Actor/Actress's Picture"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   840
      TabIndex        =   27
      Top             =   120
      Width           =   11895
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   5760
      TabIndex        =   26
      Top             =   7920
      Width           =   375
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   360
      TabIndex        =   25
      Top             =   7920
      Width           =   375
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   5760
      TabIndex        =   24
      Top             =   5640
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   360
      TabIndex        =   23
      Top             =   5640
      Width           =   375
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   5760
      TabIndex        =   22
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   5760
      TabIndex        =   21
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   360
      TabIndex        =   20
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   360
      TabIndex        =   19
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Results"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   11520
      TabIndex        =   11
      Top             =   840
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   11520
      Left            =   -1200
      Picture         =   "SpiderQuiz.frx":0000
      Top             =   0
      Width           =   15360
   End
End
Attribute VB_Name = "Spider_Quiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Ben As String, Oct As String, Ed As String, Harry As String, JJ As String, MJ As String, May As String
Private Sub Results_Click()
'set variables
Ben = Ans1.Text
Oct = Ans2.Text
Ed = Ans3.Text
Harry = Ans5.Text
Gwen = Ans4.Text
JJ = Ans6.Text
MJ = Ans7.Text
May = Ans8.Text
'answers right and wrong with printing results
If Ben = "Ben Parker" Then
    PicResults.Print "1) Correct"
Else: PicResults.Print "The correct answer for number 1"; Chr(13); "was Ben Parker"
End If
If Oct = "Dr. Otto Octavius" Then
    PicResults.Print "2) Correct"
Else: PicResults.Print "The correct answer for number 2"; Chr(13); "was Dr. Otto Octavius"
End If
If Ed = "Eddie Brock" Then
    PicResults.Print "3) Correct"
Else: PicResults.Print "The correct answer for number 3"; Chr(13); "was Eddie Brock"
End If
If Gwen = "Gwen Stacy" Then
    PicResults.Print "4) Correct"
Else: PicResults.Print "The correct answer for number 4"; Chr(13); "was Gwen Stacy"
End If
If Harry = "Harry Osborn" Then
    PicResults.Print "5) Correct"
Else: PicResults.Print "The correct answer for number 5"; Chr(13); "was Harry Osborn"
End If
If JJ = "J. Jonah Jameson" Then
    PicResults.Print "6) Correct"
Else: PicResults.Print "The correct answer for number 6"; Chr(13); "was J. Jonah Jameson"
End If
If MJ = "Mary Jane Watson" Then
    PicResults.Print "7) Correct"
Else: PicResults.Print "The correct answer for number 7"; Chr(13); "was Mary Jane Watson"
End If
If May = "May Parker" Then
    PicResults.Print "8) Correct"
Else: PicResults.Print "The correct answer for number 8"; Chr(13); "was May Parker"
End If

If Ben = "Ben Parker" And Oct = "Dr. Otto Octavius" And Ed = "Eddie Brock" And Gwen = "Gwen Stacy" And Harry = "Harry Osborn" And JJ = "J. Jonah Jameson" And MJ = "Mary Jane Watson" And May = "May Parker" Then
    PicResults.Print "Yay! Way to go "; UserName; ". All of them "; Chr(13); "are correct and"; Chr(13); "spelled correctly!"
End If
End Sub

Private Sub Show1_Click()
'show all pictures
PicResult1.Picture = LoadPicture(App.Path & "\BenParker.jpg")
PicResult2.Picture = LoadPicture(App.Path & "\DocOcttavious.jpg")
PicResult3.Picture = LoadPicture(App.Path & "\Eddie Brock.jpg")
PicResult4.Picture = LoadPicture(App.Path & "\Gwen Stacy.jpg")
PicResult5.Picture = LoadPicture(App.Path & "\HarryOsborne.jpg")
PicResult6.Picture = LoadPicture(App.Path & "\J Jonah Jameson.jpg")
PicResult7.Picture = LoadPicture(App.Path & "\Mary Jane Wattson.jpg")
PicResult8.Picture = LoadPicture(App.Path & "\MayParker.jpg")
End Sub

Private Sub Instructions_Click()
'open msgbox for instructions
MsgBox ("Click the Start button to reveal all pictures. Then Type which Character you think each actor/actress played. When you have filled in all of the blanks click Show Results in the lower left to see how you did! Don't forget to type the first and last name in that order with proper caps!")
End Sub

Private Sub MainReturn_Click()
MainMenu.Show
Spider_Quiz.Hide
End Sub

Private Sub SpiderReturn_Click()
Spiderman.Show
Spider_Quiz.Hide
End Sub
