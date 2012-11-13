VERSION 5.00
Begin VB.Form Szechuan_cuisine
   Caption         =   "Form1"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12630
   LinkTopic       =   "Form1"
   Picture         =   "Szechuan_cuisine.frx":0000
   ScaleHeight     =   9180
   ScaleWidth      =   12630
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdreturn
      BackColor       =   &H00FF80FF&
      Caption         =   "Return"
      BeginProperty Font
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10920
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8280
      Width           =   1695
   End
   Begin VB.CommandButton Command1
      BackColor       =   &H000000C0&
      Caption         =   "Let's Rock"
      BeginProperty Font
         Name            =   "Old English Text MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7920
      Width           =   1815
   End
   Begin VB.OptionButton optbomb
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   7680
      Picture         =   "Szechuan_cuisine.frx":3689E
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6600
      Width           =   855
   End
   Begin VB.OptionButton Optfire
      BackColor       =   &H0080C0FF&
      ForeColor       =   &H00C000C0&
      Height          =   855
      Left            =   6120
      Picture         =   "Szechuan_cuisine.frx":37498
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6600
      Width           =   855
   End
   Begin VB.OptionButton optsweet
      BackColor       =   &H00C0C0FF&
      ForeColor       =   &H00800000&
      Height          =   735
      Left            =   4200
      Picture         =   "Szechuan_cuisine.frx":376C0
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6600
      Width           =   975
   End
   Begin VB.PictureBox picResults
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      Height          =   4335
      Left            =   3360
      ScaleHeight     =   4275
      ScaleWidth      =   5955
      TabIndex        =   1
      Top             =   120
      Width           =   6015
   End
   Begin VB.CommandButton cmdintroduce
      BackColor       =   &H00C0C0FF&
      Caption         =   "What's Szechuan Cuisine"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   720
      MaskColor       =   &H00C0C0FF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label labquestion
      BackColor       =   &H008080FF&
      Caption         =   "how much spicy you can handle? click one of these icon and hit the button"
      BeginProperty Font
         Name            =   "Myriad Condensed Web"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   960
      TabIndex        =   5
      Top             =   6000
      Width           =   1815
   End
End
Attribute VB_Name = "Szechuan_cuisine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Chinese food
'Form Name: Szechuan
'Author: Yifan Pang
'Date Written: feb 23 2010
'The purpose of this form is introduce the szeChuan food
Option Explicit


Private Sub cmdReturn_Click()
    Szechuan_cuisine.Hide
    China.Show
End Sub

Private Sub Command1_Click() 'use option button to show picture
If optsweet = True Then
picresults.Picture = LoadPicture(App.Path & "\tanghua.jpg")
End If
If Optfire = True Then
picresults.Picture = LoadPicture(App.Path & "\dandan.jpg")
End If
If optbomb = True Then
picresults.Picture = LoadPicture(App.Path & "\huoguo.jpg")
End If


End Sub

Private Sub cmdintroduce_Click() 'load file
Dim sichuan(1 To 100) As String
Dim ctr As Integer
Dim n As Integer
Dim abc As String
picresults.Cls
Open App.Path & "\sichuan.txt" For Input As #1
    Do Until EOF(1)
        ctr = ctr + 1
        Input #1, sichuan(ctr)
    Loop
    Close #1
For n = 1 To ctr
    picresults.ForeColor = RGB(200, 0, 0)
    picresults.Print sichuan(n)
Next n
End Sub

