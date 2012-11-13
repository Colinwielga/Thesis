VERSION 5.00
Begin VB.Form Form11 
   BackColor       =   &H00FF8080&
   Caption         =   "Form11"
   ClientHeight    =   7680
   ClientLeft      =   285
   ClientTop       =   660
   ClientWidth     =   12120
   LinkTopic       =   "Form11"
   ScaleHeight     =   7680
   ScaleWidth      =   12120
   Begin VB.CommandButton cmdIdeaList 
      Caption         =   "Print a List of Programming Ideas By Subject"
      BeginProperty Font 
         Name            =   "PaintStroke"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   3495
   End
   Begin VB.PictureBox pbxResults 
      BackColor       =   &H00C0FFC0&
      FillColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "PaintStroke"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   6375
      Left            =   3840
      ScaleHeight     =   6315
      ScaleWidth      =   7995
      TabIndex        =   3
      Top             =   1200
      Width           =   8055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Back to Programming"
      BeginProperty Font 
         Name            =   "PaintStroke"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      TabIndex        =   2
      Top             =   6240
      Width           =   1695
   End
   Begin VB.CommandButton cmdForm2 
      Caption         =   "Back to Main Menu"
      BeginProperty Font 
         Name            =   "PaintStroke"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2040
      TabIndex        =   1
      Top             =   6240
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF8080&
      Caption         =   $"Form11.frx":0000
      BeginProperty Font 
         Name            =   "PaintStroke"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   3375
      Left            =   240
      TabIndex        =   5
      Top             =   2640
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Caption         =   "Active Programming"
      BeginProperty Font 
         Name            =   "PaintStroke"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   1215
      Left            =   2640
      TabIndex        =   0
      Top             =   0
      Width           =   6375
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim S As String
Dim Social(1 To 19) As String
Dim Educational(1 To 5)
Dim Wellness(1 To 14) As String
Dim Development(1 To 7) As String
Dim Careers(1 To 3) As String
Dim Intellectual(1 To 9) As String
Dim Diversity(1 To 4) As String
Dim Values(1 To 6) As String
Dim i As Integer

Sub cmdForm2_Click()
Form11.Hide
Form2.Show
End Sub

Private Sub cmdIdeaList_Click()
S = InputBox("Please enter one of the programming categories here.", "Programming")
If S = "Social" Then
Open strPath & "Social.txt" For Input As #1
For i = 1 To 19
    Input #1, Social(i)
    pbxResults.Print Social(i)
    Next i
    Close #1
ElseIf S = "Educational" Then
    Open strPath & "Educational.txt" For Input As #1
    For i = 1 To 5
    Input #1, Educational(i)
    pbxResults.Print Educational(i)
    Next i
    Close #1
ElseIf S = "Health And Wellness" Then
    Open strPath & "HealthWellness.txt" For Input As #1
    For i = 1 To 14
    Input #1, Wellness(i)
    pbxResults.Print Wellness(i)
    Next i
    Close #1
ElseIf S = "Women's Development" Then
    Open strPath & "Women'sDevelopment.txt" For Input As #1
    For i = 1 To 7
    Input #1, Development(i)
    pbxResults.Print Development(i)
    Next i
    Close #1
ElseIf S = "Careers" Then
    Open strPath & "Careers.txt" For Input As #1
    For i = 1 To 3
    Input #1, Careers(i)
    pbxResults.Print Careers(i)
    Next i
    Close #1
ElseIf S = "Intellectual" Then
    Open strPath & "Intellectual.txt" For Input As #1
    For i = 1 To 9
    Input #1, Intellectual(i)
    pbxResults.Print Intellectual(i)
    Next i
    Close #1
ElseIf S = "Diversity" Then
    Open strPath & "Diversity.txt" For Input As #1
    For i = 1 To 4
    Input #1, Diversity(i)
    pbxResults.Print Diversity(i)
    Next i
    Close #1
ElseIf S = "Benedictine Values/Spirituality" Then
    Open strPath & "BenedictineValues.txt" For Input As #1
    For i = 1 To 6
    Input #1, Values(i)
    pbxResults.Print Values(i)
    Next i
    Close #1
End If
Close #1
End Sub

Private Sub Command1_Click()
Form11.Hide
Form4.Show
End Sub

