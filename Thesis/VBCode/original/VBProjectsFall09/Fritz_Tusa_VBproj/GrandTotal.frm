VERSION 5.00
Begin VB.Form GrandTotal 
   BackColor       =   &H80000007&
   Caption         =   "GrandTotal"
   ClientHeight    =   9555
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13845
   LinkTopic       =   "Form1"
   ScaleHeight     =   9555
   ScaleWidth      =   13845
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF8080&
      Caption         =   $"GrandTotal.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6600
      Width           =   2895
   End
   Begin VB.PictureBox picResultsTotal 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   3480
      ScaleHeight     =   3795
      ScaleWidth      =   7155
      TabIndex        =   2
      Top             =   480
      Width           =   7215
   End
   Begin VB.CommandButton cmdshogrand 
      BackColor       =   &H00FF00FF&
      Caption         =   "SHOW YOUR GRAND TOTAL!"
      BeginProperty Font 
         Name            =   "Bauhaus 93"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4440
      Width           =   3735
   End
   Begin VB.CommandButton cmdtitlael 
      BackColor       =   &H008080FF&
      Caption         =   "To Title"
      BeginProperty Font 
         Name            =   "Jokerman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   11880
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8040
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   2145
      Left            =   8760
      Picture         =   "GrandTotal.frx":00A8
      Stretch         =   -1  'True
      Top             =   7200
      Width           =   1725
   End
End
Attribute VB_Name = "GrandTotal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'SKI TRIP'
'GRAND TOTAL'
'MAX TUSA'
'8-18'
'THIS FORM DISPLAYS THE GRAND TOTAL FROM ALL FORMS'

Private Sub cmdshogrand_Click()
'clear the picture box before beginning'
picResultsTotal.Cls


totalCost = skiticketcost + skirentalcost + totalairfarecost + totalhotelcost

'show the total money to be spent'
picResultsTotal.Print "The cost for your trip to "; resorts(place); " would be "; FormatCurrency(totalCost)
picResultsTotal.Print ""
picResultsTotal.Print "Your Total was calculated on"
picResultsTotal.Print Date; "at approximately"; Time
picResultsTotal.Print ""
picResultsTotal.Print "Print this off and run around,"; Tab(1); "showing everyone you know how"; Tab(1); "much your awesome trip will cost!!!!"

Open App.Path & "\YOURCOST.txt" For Output As #1

    Print #1, "The cost for your trip to "; resorts(place); " would be "; FormatCurrency(totalCost)
    Print #1, ""
    Print #1, "Your Total was calculated on"
    Print #1, Date; "at approximately  "; Time
    Print #1, ""
    Print #1, "Print this off and run around,"; Tab(1); "showing everyone you know how"; Tab(1); "much your awesome trip will cost!!!!"

Close #1


End Sub

Private Sub cmdtitlael_Click()
Title.Show
GrandTotal.Hide
End Sub

Private Sub Command1_Click()
Dim linne As String, temp As String
temp = App.Path
linne = "notepad " & temp & "\YOURCOST.txt"
Shell linne
End Sub

