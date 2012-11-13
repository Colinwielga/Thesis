VERSION 5.00
Begin VB.Form VolkswagonForm 
   BackColor       =   &H80000010&
   Caption         =   "Wolkswagon"
   ClientHeight    =   8790
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11235
   FillColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8790
   ScaleWidth      =   11235
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdInput 
      Caption         =   "Input"
      BeginProperty Font 
         Name            =   "Juice ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   9
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CommandButton cmdMainForm1 
      BackColor       =   &H000000FF&
      Caption         =   "Back to Main Form 1"
      BeginProperty Font 
         Name            =   "Juice ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9480
      TabIndex        =   8
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmdMake 
      BackColor       =   &H000000FF&
      Caption         =   "Back to Make Form"
      BeginProperty Font 
         Name            =   "Juice ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7920
      TabIndex        =   7
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmdFacts 
      BackColor       =   &H000000FF&
      Caption         =   "Volkswagon Quotes"
      BeginProperty Font 
         Name            =   "Juice ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8400
      TabIndex        =   6
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdPush 
      BackColor       =   &H000000FF&
      Caption         =   "Push"
      BeginProperty Font 
         Name            =   "Juice ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   5640
      Width           =   855
   End
   Begin VB.PictureBox picOutput 
      BackColor       =   &H80000014&
      Height          =   2655
      Left            =   2880
      ScaleHeight     =   2595
      ScaleWidth      =   4635
      TabIndex        =   3
      Top             =   4080
      Width           =   4695
   End
   Begin VB.Image Image12 
      Height          =   720
      Left            =   9480
      Picture         =   "Volkswagon.frx":0000
      Top             =   7680
      Width           =   1140
   End
   Begin VB.Image Image11 
      Height          =   675
      Left            =   8280
      Picture         =   "Volkswagon.frx":07B6
      Top             =   7680
      Width           =   1065
   End
   Begin VB.Image Image10 
      Height          =   855
      Left            =   6960
      Picture         =   "Volkswagon.frx":0EDA
      Top             =   7560
      Width           =   1140
   End
   Begin VB.Image Image9 
      Height          =   675
      Left            =   5640
      Picture         =   "Volkswagon.frx":2EA7
      Top             =   7680
      Width           =   1095
   End
   Begin VB.Image Image8 
      Height          =   720
      Left            =   4320
      Picture         =   "Volkswagon.frx":3339
      Top             =   7680
      Width           =   1140
   End
   Begin VB.Image Image7 
      Height          =   675
      Left            =   3000
      Picture         =   "Volkswagon.frx":3AEF
      Top             =   7680
      Width           =   1065
   End
   Begin VB.Image Image6 
      Height          =   855
      Left            =   1560
      Picture         =   "Volkswagon.frx":4213
      Top             =   7560
      Width           =   1140
   End
   Begin VB.Image Image5 
      Height          =   675
      Left            =   120
      Picture         =   "Volkswagon.frx":61E0
      Top             =   7680
      Width           =   1095
   End
   Begin VB.Image Image4 
      Height          =   720
      Left            =   9360
      Picture         =   "Volkswagon.frx":6672
      Top             =   120
      Width           =   1140
   End
   Begin VB.Image Image3 
      Height          =   675
      Left            =   8040
      Picture         =   "Volkswagon.frx":6E28
      Top             =   120
      Width           =   1065
   End
   Begin VB.Image Image2 
      Height          =   855
      Left            =   6720
      Picture         =   "Volkswagon.frx":754C
      Top             =   120
      Width           =   1140
   End
   Begin VB.Image Image1 
      Height          =   675
      Left            =   5400
      Picture         =   "Volkswagon.frx":9519
      Top             =   120
      Width           =   1095
   End
   Begin VB.Image Convertible 
      Height          =   720
      Left            =   4080
      Picture         =   "Volkswagon.frx":99AB
      Top             =   120
      Width           =   1140
   End
   Begin VB.Image SUV 
      Height          =   675
      Left            =   2760
      Picture         =   "Volkswagon.frx":A161
      Top             =   120
      Width           =   1065
   End
   Begin VB.Image sedan 
      Height          =   855
      Left            =   1440
      Picture         =   "Volkswagon.frx":A885
      Top             =   120
      Width           =   1140
   End
   Begin VB.Image Truck 
      Height          =   675
      Left            =   120
      Picture         =   "Volkswagon.frx":C852
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblOutput 
      BackColor       =   &H0080FFFF&
      Caption         =   "Here are some of the cars that match your requirements:"
      BeginProperty Font 
         Name            =   "Juice ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1320
      TabIndex        =   4
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label lblstar 
      Caption         =   "****************************************************"
      BeginProperty Font 
         Name            =   "Juice ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   3960
      TabIndex        =   2
      Top             =   2520
      Width           =   3135
   End
   Begin VB.Label lblStyles 
      BackColor       =   &H0080FFFF&
      Caption         =   "Trucks  Sedans  SUVs  Convertibles"
      BeginProperty Font 
         Name            =   "Juice ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   2760
      Width           =   3015
   End
   Begin VB.Label lblStyle 
      BackColor       =   &H0080FFFF&
      Caption         =   "Which style of car are you interested in?"
      BeginProperty Font 
         Name            =   "Juice ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   0
      Top             =   1920
      Width           =   2295
   End
End
Attribute VB_Name = "VolkswagonForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Form Name : VolkswagonForm(Volkwagon.frm)
'Author : Katie Lee
'Purpose : Asks user to input a specific make/style and displays the model and pri
Option Explicit
Dim Make As String

Private Sub cmdInput_Click()
Dim Style As String
Make = InputBox("Please enter the style of car in which you are interested in", "Style of car") 'Asks user to input make/style
picOutput.Print "You are interested in looking at a:"; Tab(1); Make
End Sub

Private Sub cmdMainForm1_Click()
VolkswagonForm.Hide
MainForm1.Show ' Bring user back to MainForm1

End Sub

Private Sub cmdMake_Click()
VolkswagonForm.Hide
MakeForm.Show ' Bring user MakeForm

End Sub

Private Sub cmdPush_Click()
Dim CTR As Integer
For CTR = 1 To 39
    If Company(CTR) = "Volkswagon" And Style(CTR) = Make Then 'Searches for Volkswagon companies and style that matches the input
        picOutput.Print "Volkswagon", Model(CTR), FormatCurrency(Price(CTR))
    End If
Next CTR
End Sub
