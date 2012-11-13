VERSION 5.00
Begin VB.Form frmConversion 
   Caption         =   "Pace Calculator"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      Height          =   4575
      Left            =   4800
      Picture         =   "frmconversion.frx":0000
      ScaleHeight     =   4515
      ScaleWidth      =   3915
      TabIndex        =   15
      Top             =   5040
      Width           =   3975
   End
   Begin VB.PictureBox Picture1 
      Height          =   5415
      Left            =   9120
      Picture         =   "frmconversion.frx":41EF2
      ScaleHeight     =   5355
      ScaleWidth      =   5235
      TabIndex        =   14
      Top             =   2040
      Width           =   5295
   End
   Begin VB.CommandButton cmddir 
      Caption         =   "Back To Directory"
      Height          =   975
      Left            =   10440
      TabIndex        =   13
      Top             =   7920
      Width           =   2535
   End
   Begin VB.CommandButton cdmconversion 
      Caption         =   "Convert"
      Height          =   735
      Left            =   1320
      TabIndex        =   12
      Top             =   6840
      Width           =   2175
   End
   Begin VB.TextBox txtmtime 
      Height          =   525
      Left            =   2880
      TabIndex        =   11
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton cmdcompute 
      Caption         =   "Convert"
      Height          =   735
      Left            =   1200
      TabIndex        =   9
      Top             =   3120
      Width           =   2175
   End
   Begin VB.PictureBox picresults 
      Height          =   1215
      Left            =   4800
      ScaleHeight     =   1155
      ScaleWidth      =   3915
      TabIndex        =   8
      Top             =   3360
      Width           =   3975
   End
   Begin VB.TextBox txttime 
      Height          =   495
      Left            =   3000
      TabIndex        =   7
      Top             =   2400
      Width           =   1455
   End
   Begin VB.TextBox txtmile 
      Height          =   495
      Left            =   2880
      TabIndex        =   5
      Top             =   4920
      Width           =   1455
   End
   Begin VB.TextBox txtkilo 
      Height          =   405
      Left            =   3000
      TabIndex        =   2
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label9 
      Caption         =   "November 5, 2008"
      Height          =   255
      Left            =   2520
      TabIndex        =   19
      Top             =   10320
      Width           =   2895
   End
   Begin VB.Label Label8 
      Caption         =   "By: Tyler Trettel and Josh Gunderson"
      Height          =   255
      Left            =   2520
      TabIndex        =   18
      Top             =   10080
      Width           =   2895
   End
   Begin VB.Label Label7 
      Caption         =   "Pace Calculator"
      Height          =   255
      Left            =   0
      TabIndex        =   17
      Top             =   10320
      Width           =   2535
   End
   Begin VB.Label Label6 
      Caption         =   "2008 MIAC Cross Country Project "
      Height          =   255
      Left            =   0
      TabIndex        =   16
      Top             =   10080
      Width           =   2535
   End
   Begin VB.Label Label5 
      Caption         =   "How long did it take you? (Round to the Nearest Minute)"
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   5880
      Width           =   2175
   End
   Begin VB.Label Label4 
      Caption         =   "How long did it take you? (Round to the Nearest Minute)"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "How many miles did you run?"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   5040
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Please enter how many Miles or Kilometers you ran along with how long it took you."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      Top             =   1080
      Width           =   8175
   End
   Begin VB.Label lblkilo 
      Caption         =   "How many  Kilometers did you run?"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Pace Calculator"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   5400
      TabIndex        =   0
      Top             =   240
      Width           =   5295
   End
End
Attribute VB_Name = "frmconversion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Kilo As Single
Dim Mile As Single
Dim Time As Integer
Dim mtime As Integer
Dim Div As Single
Dim seconds As Single
Dim time1 As Integer
'Project Name: MIAC CC Project
'Form Name: frmConversion
'Authors: Josh Gunderson & Tyler Trettel
'Date: 5 November 2008
'Objective: The purpose of this form is for the user to convert miles to kilometers and vice versa.  Along with this form the  user is able to calculate their pace for a distance they ran







Private Sub cdmconversion_Click()
picresults.Cls
Dim Y As Single
Dim Pace1 As Single
Mile = txtmile
time1 = txtmtime
Dim Second As Single
Dim D As Integer
Dim Divi As Single

Y = Mile * 1.61
Divi = time1 / Mile
Second = Divi * 60

Do While Second > 59
    Second = Second - 60
    D = D + 1
Loop



picresults.Print Mile; "Miles is equivilent to"; Y; "kilometers."
picresults.Print ""
picresults.Print "Congrats your pace was"; D; ":"; Second; " Per Mile."



End Sub


Private Sub cmdcompute_Click()
picresults.Cls
Dim X As Single
Dim Pace As Single
Dim G As Integer
Kilo = txtkilo
Time = txttime

X = Kilo * 0.62
Div = Time / Kilo

seconds = Div * 60

Do While seconds > 59
    seconds = seconds - 60
    G = G + 1
Loop



picresults.Print Kilo; "Kilometers is equivilent to"; X; "miles."
picresults.Print ""
picresults.Print "Congrats your pace was"; G; ":"; FormatNumber(seconds, 0); " Per Kilometer."




End Sub

Private Sub cmddir_Click()
frmconversion.Hide
frmdirectory.Show
End Sub


