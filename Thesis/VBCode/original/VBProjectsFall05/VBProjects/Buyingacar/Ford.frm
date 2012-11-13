VERSION 5.00
Begin VB.Form FordForm 
   BackColor       =   &H80000011&
   Caption         =   "Ford"
   ClientHeight    =   8970
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10695
   LinkTopic       =   "Form1"
   ScaleHeight     =   8970
   ScaleWidth      =   10695
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
      Height          =   735
      Left            =   4200
      TabIndex        =   9
      Top             =   3240
      Width           =   1935
   End
   Begin VB.CommandButton cmdMake 
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
      TabIndex        =   8
      Top             =   6720
      Width           =   975
   End
   Begin VB.CommandButton cmdMainForm1 
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
      Left            =   9120
      TabIndex        =   7
      Top             =   6720
      Width           =   1095
   End
   Begin VB.CommandButton cmdFacts 
      Caption         =   "Ford Quotes"
      BeginProperty Font 
         Name            =   "Juice ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8160
      TabIndex        =   6
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton cmsPush 
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
      Height          =   495
      Left            =   1560
      TabIndex        =   5
      Top             =   5640
      Width           =   855
   End
   Begin VB.PictureBox picOutput 
      Height          =   3375
      Left            =   3000
      ScaleHeight     =   3315
      ScaleWidth      =   4635
      TabIndex        =   3
      Top             =   4080
      Width           =   4695
   End
   Begin VB.Image Image12 
      Height          =   720
      Left            =   8040
      Picture         =   "Ford.frx":0000
      Top             =   8040
      Width           =   1140
   End
   Begin VB.Image Image11 
      Height          =   675
      Left            =   9360
      Picture         =   "Ford.frx":07B6
      Top             =   8040
      Width           =   1065
   End
   Begin VB.Image Image10 
      Height          =   855
      Left            =   6720
      Picture         =   "Ford.frx":0EDA
      Top             =   7920
      Width           =   1140
   End
   Begin VB.Image Image9 
      Height          =   675
      Left            =   5400
      Picture         =   "Ford.frx":2EA7
      Top             =   8040
      Width           =   1095
   End
   Begin VB.Image Image8 
      Height          =   720
      Left            =   4080
      Picture         =   "Ford.frx":3339
      Top             =   8040
      Width           =   1140
   End
   Begin VB.Image Image7 
      Height          =   675
      Left            =   2760
      Picture         =   "Ford.frx":3AEF
      Top             =   8040
      Width           =   1065
   End
   Begin VB.Image Image6 
      Height          =   855
      Left            =   1440
      Picture         =   "Ford.frx":4213
      Top             =   7920
      Width           =   1140
   End
   Begin VB.Image Image5 
      Height          =   675
      Left            =   120
      Picture         =   "Ford.frx":61E0
      Top             =   8040
      Width           =   1095
   End
   Begin VB.Image Image4 
      Height          =   720
      Left            =   9360
      Picture         =   "Ford.frx":6672
      Top             =   240
      Width           =   1140
   End
   Begin VB.Image Image3 
      Height          =   675
      Left            =   8040
      Picture         =   "Ford.frx":6E28
      Top             =   240
      Width           =   1065
   End
   Begin VB.Image Image2 
      Height          =   855
      Left            =   6720
      Picture         =   "Ford.frx":754C
      Top             =   120
      Width           =   1140
   End
   Begin VB.Image Image1 
      Height          =   675
      Left            =   5400
      Picture         =   "Ford.frx":9519
      Top             =   240
      Width           =   1095
   End
   Begin VB.Image Convertible 
      Height          =   720
      Left            =   4080
      Picture         =   "Ford.frx":99AB
      Top             =   240
      Width           =   1140
   End
   Begin VB.Image SUV 
      Height          =   675
      Left            =   2760
      Picture         =   "Ford.frx":A161
      Top             =   240
      Width           =   1065
   End
   Begin VB.Image sedan 
      Height          =   855
      Left            =   1440
      Picture         =   "Ford.frx":A885
      Top             =   120
      Width           =   1140
   End
   Begin VB.Image Truck 
      Height          =   675
      Left            =   120
      Picture         =   "Ford.frx":C852
      Top             =   240
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
      Height          =   1215
      Left            =   1320
      TabIndex        =   4
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label lblstar 
      Caption         =   "************************************************"
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
      Left            =   3600
      TabIndex        =   2
      Top             =   2400
      Width           =   3015
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
      Left            =   3600
      TabIndex        =   1
      Top             =   2640
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
      Height          =   615
      Left            =   3960
      TabIndex        =   0
      Top             =   1680
      Width           =   2295
   End
End
Attribute VB_Name = "FordForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : Buying A Car(VB-project.vbp)
'Form Name : Ford(Ford.frm)
'Author: Katie Lee
'purpose of the form:  Since the user has selected a Brand, the purpose
                    'of this form is for the user to determine which style of car
                    'he/she needs depending on their preference
Option Explicit
Dim Make As String
'Dim Style(1 To 39) As Integer

Private Sub cmdFacts_Click()
MsgBox " Ford Earns Top Ratings in 2003 Frontal Crash Tests in the models: the Ford Expedition 4x4, Ford Crown Victoria, Ford Windstar, Ford Taurus, Ford Mustang 2-door, Ford F-150 Supercrew 4-door 4x2.  The Ford F-150 has earned the “Highest Ranked” spot in the full-size light-duty pickup segment of the J.D. Power and Associates 2005 Initial Quality StudySM, (competitor information cannot be used).", , "Ford Ratings"
    'brings up message box to  show the  user quotes pertaining to Ford
End Sub

Private Sub cmdInput_Click()

Make = InputBox("Please enter the style of car in which you are interested in", "Style of car") 'gets the car style from the user
picOutput.Print Make

End Sub

Private Sub cmdMainForm1_Click()
FordForm.Hide
MainForm1.Show 'Returns user to MainForm1

End Sub

Private Sub cmdMake_Click()
FordForm.Hide
MakeForm.Show 'Returns User to MakeForm

End Sub

Private Sub cmsPush_Click()
Dim CTR As Integer
For CTR = 1 To 39
    If Company(CTR) = "Ford" And Style(CTR) = Make Then 'Searches for Ford companies and style that matches the input
        picOutput.Print "Ford", Model(CTR), FormatCurrency(Price(CTR))
    End If
Next CTR

End Sub

