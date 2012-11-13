VERSION 5.00
Begin VB.Form ToyotaForm 
   BackColor       =   &H80000010&
   Caption         =   "Toyota"
   ClientHeight    =   8685
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11010
   LinkTopic       =   "Form1"
   ScaleHeight     =   8685
   ScaleWidth      =   11010
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
      Height          =   615
      Left            =   4440
      TabIndex        =   9
      Top             =   3720
      Width           =   1575
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
      Left            =   9360
      TabIndex        =   8
      Top             =   6480
      Width           =   1095
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
      TabIndex        =   7
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CommandButton cmdFacts 
      Caption         =   "Toyota Quotes"
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
      Left            =   8400
      TabIndex        =   6
      Top             =   4560
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
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   6480
      Width           =   855
   End
   Begin VB.PictureBox picOutput 
      BackColor       =   &H80000014&
      Height          =   2415
      Left            =   3000
      ScaleHeight     =   2355
      ScaleWidth      =   4515
      TabIndex        =   3
      Top             =   4560
      Width           =   4575
   End
   Begin VB.Image Image12 
      Height          =   720
      Left            =   9480
      Picture         =   "Toyota.frx":0000
      Top             =   7680
      Width           =   1140
   End
   Begin VB.Image Image11 
      Height          =   675
      Left            =   8160
      Picture         =   "Toyota.frx":07B6
      Top             =   7680
      Width           =   1065
   End
   Begin VB.Image Image10 
      Height          =   855
      Left            =   6840
      Picture         =   "Toyota.frx":0EDA
      Top             =   7560
      Width           =   1140
   End
   Begin VB.Image Image9 
      Height          =   675
      Left            =   5520
      Picture         =   "Toyota.frx":2EA7
      Top             =   7680
      Width           =   1095
   End
   Begin VB.Image Image8 
      Height          =   720
      Left            =   4080
      Picture         =   "Toyota.frx":3339
      Top             =   7680
      Width           =   1140
   End
   Begin VB.Image Image7 
      Height          =   675
      Left            =   2760
      Picture         =   "Toyota.frx":3AEF
      Top             =   7680
      Width           =   1065
   End
   Begin VB.Image Image6 
      Height          =   855
      Left            =   1440
      Picture         =   "Toyota.frx":4213
      Top             =   7680
      Width           =   1140
   End
   Begin VB.Image Image5 
      Height          =   675
      Left            =   120
      Picture         =   "Toyota.frx":61E0
      Top             =   7680
      Width           =   1095
   End
   Begin VB.Image Image4 
      Height          =   720
      Left            =   9360
      Picture         =   "Toyota.frx":6672
      Top             =   240
      Width           =   1140
   End
   Begin VB.Image Image3 
      Height          =   675
      Left            =   8160
      Picture         =   "Toyota.frx":6E28
      Top             =   240
      Width           =   1065
   End
   Begin VB.Image Image2 
      Height          =   855
      Left            =   6840
      Picture         =   "Toyota.frx":754C
      Top             =   120
      Width           =   1140
   End
   Begin VB.Image Image1 
      Height          =   675
      Left            =   5520
      Picture         =   "Toyota.frx":9519
      Top             =   120
      Width           =   1095
   End
   Begin VB.Image Convertible 
      Height          =   720
      Left            =   4200
      Picture         =   "Toyota.frx":99AB
      Top             =   120
      Width           =   1140
   End
   Begin VB.Image SUV 
      Height          =   675
      Left            =   2880
      Picture         =   "Toyota.frx":A161
      Top             =   120
      Width           =   1065
   End
   Begin VB.Image sedan 
      Height          =   855
      Left            =   1560
      Picture         =   "Toyota.frx":A885
      Top             =   120
      Width           =   1140
   End
   Begin VB.Image Truck 
      Height          =   675
      Left            =   240
      Picture         =   "Toyota.frx":C852
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
      Left            =   1560
      TabIndex        =   4
      Top             =   4560
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
      Left            =   3600
      TabIndex        =   2
      Top             =   2880
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
      Left            =   3720
      TabIndex        =   1
      Top             =   3120
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
      Left            =   4080
      TabIndex        =   0
      Top             =   2280
      Width           =   2295
   End
End
Attribute VB_Name = "ToyotaForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : Buying A Car(VB-project.vbp)
'Form Name : ToyotaForm(Toyota.frm)
'Author : Katie Lee
'Purpose : Asks user to input a specific make/style and displays the model and price
Option Explicit
Dim Make As String

Private Sub cmdFacts_Click()
MsgBox "Toyota has more vehicles rated most fuel-efficient in their classes than any other automotive brand. Based on 2001 Environmental Protection Agency (EPA) Most Fuel Efficient rankings, the ECHO subcompact sedan, RAV4 compact sport utility vehicle (SUV), Avalon full-size sedan, Tacoma compact pickup, Prius gasoline/electric hybrid sedan are all leaders in their respective classes for fuel economy.", , "Toyota Quotes"
    'brings up message box to show the user quotes pertaining to Toyota
End Sub

Private Sub cmdInput_Click()

Make = InputBox("Please enter the style of car in which you are interested in", "Style of car") 'Asks user to input make/style
picOutput.Print "You are interested in looking at a:"; Tab(1); Make
End Sub

Private Sub cmdMainForm1_Click()
ToyotaForm.Hide
MainForm1.Show 'Sends user back to MainForm1

End Sub

Private Sub cmdMake_Click()
ToyotaForm.Hide
MakeForm.Show 'Sends user to MakeForm

End Sub

Private Sub cmsPush_Click()
picOutput.Print "*******************************************************"
Dim CTR As Integer
For CTR = 1 To 39
    If Company(CTR) = "Toyota" And Style(CTR) = Make Then 'Searches for Toyota companies and style that matches the input
        picOutput.Print "Toyota", Model(CTR), FormatCurrency(Price(CTR))
    
    End If
Next CTR
End Sub
