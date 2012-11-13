VERSION 5.00
Begin VB.Form Overview 
   Caption         =   "Overview of Chinese Zodiacs"
   ClientHeight    =   9555
   ClientLeft      =   3735
   ClientTop       =   1290
   ClientWidth     =   12780
   FillColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   Picture         =   "Overview.frx":0000
   ScaleHeight     =   9555
   ScaleWidth      =   12780
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H80000009&
      Caption         =   "Back to Main"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   9120
      MaskColor       =   &H00FFC0C0&
      Picture         =   "Overview.frx":D6B2E
      TabIndex        =   4
      Top             =   360
      Width           =   2415
   End
   Begin VB.CommandButton cmdSort2 
      Caption         =   "Put the Zodiacs in ascending alphabetical orders"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   6240
      Picture         =   "Overview.frx":D7FDD
      TabIndex        =   3
      Top             =   360
      Width           =   2535
   End
   Begin VB.CommandButton cmdSort 
      Caption         =   "Put the Zodiacs in ascending numerical orders"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3600
      TabIndex        =   2
      Top             =   360
      Width           =   2415
   End
   Begin VB.CommandButton cmdOverview 
      Caption         =   "Chinese Zodiac Overview"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   960
      TabIndex        =   1
      Top             =   360
      Width           =   2415
   End
   Begin VB.PictureBox picResults 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   7425
      Left            =   960
      Picture         =   "Overview.frx":D948C
      ScaleHeight     =   7365
      ScaleWidth      =   10515
      TabIndex        =   0
      Top             =   1680
      Width           =   10575
   End
End
Attribute VB_Name = "Overview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdBack_Click()
Dim TempZodiac As String, TempNum As Integer, I As Integer, J As Integer
overview.Visible = False
Home.Visible = True
 For I = 1 To 11                                   'This is to make sure, after the user has used the alphabetical sorting,
            For J = 1 To 12 - I                    'the arrays of zodiacs and their names are still in the original order.
            If num(J) > num(J + 1) Then
                TempZodiac = zodiac(J)
                zodiac(J) = zodiac(J + 1)
                zodiac(J + 1) = TempZodiac
                TempNum = num(J)
                num(J) = num(J + 1)
                num(J + 1) = TempNum
            End If
            Next J
    Next I
MsgBox "Welcome back! What else can I do for you?", , "Welcome!"
End Sub

Private Sub cmdOverview_Click()
Dim overview(1 To 100) As String, I As Integer, Ctr As Integer
picResults.Cls
picResults.Picture = LoadPicture(App.Path & "\images\" & "2010 chinese new year.jpg")
picResults.ForeColor = vbWhite
picResults.Font = "Papyrus"
picResults.FontSize = 9
Open App.Path & "\overview.txt" For Input As #1    'put the comments about the zodiac in to an array of strings
    Do Until EOF(1)
    Ctr = Ctr + 1
    Input #1, overview(Ctr)
    Loop
Close #1
For I = 1 To Ctr                                  'out put all of the strings
picResults.Print overview(I)
Next I
End Sub

Private Sub cmdSort_Click()
Dim TempZodiac As String, TempNum As Integer, I As Integer, J As Integer
picResults.Cls
picResults.Picture = LoadPicture(App.Path & "\images\" & "The year of Tiger.jpg")     'change a background picture
picResults.Font = "Verdana"
picResults.FontSize = 16
picResults.ForeColor = vbRed                        'changing the color of the header, which is not that interesting
picResults.Print "Names of the Zodiacs", "Number"
picResults.Print "********************************************"
picResults.ForeColor = vbWhite
    For I = 1 To 11                                 'sorting numerically. This seemed useless, because when the zodiacs
            For J = 1 To 12 - I                     'were read into the file, they were in numeric orders. However,
            If num(J) > num(J + 1) Then             'After executing the the other sorting button, the order will be wrong.
                TempZodiac = zodiac(J)              'So, this process is necessary when user want to see the alphabetical order first
                zodiac(J) = zodiac(J + 1)           'and the numeric order second.
                zodiac(J + 1) = TempZodiac
                TempNum = num(J)
                num(J) = num(J + 1)
                num(J + 1) = TempNum
            End If
            Next J
    Next I
    For I = 1 To 12
        picResults.Print zodiac(I), Tab(20), num(I)
    Next I
End Sub

Private Sub cmdSort2_Click()
Dim TempZodiac As String, TempNum As Integer, I As Integer, J As Integer
picResults.Cls
picResults.Picture = LoadPicture(App.Path & "\images\" & "The year of Tiger.jpg")    'Change a backgroud picture
picResults.Font = "Verdana"
picResults.FontSize = 16
picResults.ForeColor = vbRed
picResults.Print "Names of the Zodiacs", "Number"
picResults.Print "*********************************************"
picResults.ForeColor = vbWhite
    For I = 1 To 11
        For J = 1 To 12 - I
        If zodiac(J) > zodiac(J + 1) Then        'Sorting alphabetically
            TempZodiac = zodiac(J)
            zodiac(J) = zodiac(J + 1)
            zodiac(J + 1) = TempZodiac
            TempNum = num(J)
            num(J) = num(J + 1)
            num(J + 1) = TempNum
        End If
        Next J
    Next I
    For I = 1 To 12
    picResults.Print zodiac(I), Tab(20), num(I)
    Next I

End Sub

