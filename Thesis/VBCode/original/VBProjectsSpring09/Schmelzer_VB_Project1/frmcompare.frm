VERSION 5.00
Begin VB.Form frmcompare 
   BackColor       =   &H00000000&
   Caption         =   "compare"
   ClientHeight    =   7725
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   10200
   LinkTopic       =   "Form1"
   ScaleHeight     =   7725
   ScaleWidth      =   10200
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdback 
      Caption         =   "Return to previous page"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   7
      Top             =   5760
      Width           =   1575
   End
   Begin VB.PictureBox picresults 
      BackColor       =   &H00FF0000&
      FillColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   2775
      Left            =   3360
      ScaleHeight     =   2715
      ScaleWidth      =   6315
      TabIndex        =   6
      Top             =   4200
      Width           =   6375
   End
   Begin VB.TextBox txtsecond 
      Height          =   975
      Left            =   4440
      TabIndex        =   5
      Top             =   2880
      Width           =   4095
   End
   Begin VB.TextBox txtfirst 
      Height          =   855
      Left            =   4440
      TabIndex        =   4
      Top             =   1560
      Width           =   3975
   End
   Begin VB.CommandButton cmdcompare 
      Caption         =   "click to compare"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   360
      TabIndex        =   0
      Top             =   4200
      Width           =   2535
   End
   Begin VB.Label lblquit 
      BackColor       =   &H00FF00FF&
      Caption         =   "Click above to quit"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblsecond 
      BackColor       =   &H000080FF&
      Caption         =   "Type second show                   Use lower case and make it one word "
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      TabIndex        =   3
      Top             =   2880
      Width           =   3255
   End
   Begin VB.Label lblfirst 
      BackColor       =   &H000080FF&
      Caption         =   "Type first show.                   Use lower case and make it one word"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   2
      Top             =   1560
      Width           =   3255
   End
   Begin VB.Label lblcompare 
      BackColor       =   &H0000FFFF&
      Caption         =   "Compare 2 shows to see who has the higher rating"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      TabIndex        =   1
      Top             =   600
      Width           =   6975
   End
   Begin VB.Menu file 
      Caption         =   "file"
      Begin VB.Menu Quit 
         Caption         =   "Quit"
      End
   End
End
Attribute VB_Name = "frmcompare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TV Frenzy
'Maija Schmelzer
'2/20
'this form will compare the ratings of two shows



Private Sub cmdback_Click()
'this will allow user to change forms

frmcompare.Hide
frmabout.show

End Sub

Private Sub cmdcompare_Click()
'this subroutine compares the ratings of two shows and tells the user which one has the higher rating

Dim show1 As String, show2 As String, show As String, shows(1 To 100) As String, ctr As Integer
Dim rating(1 To 100) As Single, seinfeld As String, theoffice As String
Dim scrubs As String, friends As String
Dim onetreehill As String, trustme As String, lawandorder As String, medium As String
Dim twentyfour As String, heroes As String, lost As String, savinggrace As String
Dim smallville As String, dancingwiththestars As String, realworld As String
Dim americanidol As String, thebiggestloser As String, bones As String, house As String
Dim supernatural As String, thecloser As String, greysanatomy As String, er As String
Dim truelife As String, I As Integer, found As Boolean, ctr1 As Integer, ctr2 As Integer


picresults.Print "The higher rated show of the two is..."
picresults.Print "**********************************************************************"
picresults.Print
picresults.Print
ctr = 0
Open App.Path & "\compareratings.txt" For Input As #2

Do While Not EOF(2)
    ctr = ctr + 1
    Input #2, rating(ctr), shows(ctr)
Loop
Close #2




show1 = txtfirst.Text
show2 = txtsecond.Text
ctr1 = 0
ctr2 = 0
found = False
I = 0
Do While Not found And ctr1 < ctr
    ctr1 = ctr1 + 1
    If shows(ctr1) = show1 Then
    found = True
    End If
Loop
found = False
Do While Not found And ctr2 < ctr
    ctr2 = ctr2 + 1
    If shows(ctr2) = show2 Then
    found = True
    End If
Loop



      
        
        If rating(ctr1) > rating(ctr2) Then
            picresults.Cls
            picresults.Print "The show with the higher rating is...."
            picresults.Print
            picresults.Print
            picresults.Print Tab(25); show1
        
        ElseIf rating(ctr1) < rating(ctr2) Then
            picresults.Cls
            picresults.Print "The show with the higher rating is...."
            picresults.Print
            picresults.Print
            picresults.Print Tab(25); show2
           
        ElseIf rating(ctr1) = rating(ctr2) Then
            picresults.Cls
            picresults.Print
            picresults.Print
            picresults.Print "these two shows have the same rating"
          
        End If




End Sub



Private Sub Quit_Click()
End
End Sub
