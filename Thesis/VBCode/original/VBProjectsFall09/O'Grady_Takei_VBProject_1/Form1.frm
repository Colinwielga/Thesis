VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5370
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8655
   LinkTopic       =   "Form1"
   ScaleHeight     =   5370
   ScaleWidth      =   8655
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdtotal 
      Caption         =   "Total"
      Height          =   855
      Left            =   4560
      TabIndex        =   10
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CommandButton cmdspanish200 
      Caption         =   "Spanish 200"
      Height          =   855
      Left            =   2160
      TabIndex        =   9
      Top             =   3000
      Width           =   1815
   End
   Begin VB.CommandButton cmdsoc201 
      Caption         =   "Sociology201"
      Height          =   735
      Left            =   2160
      TabIndex        =   8
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CommandButton cmdreturn 
      Caption         =   "Return"
      Height          =   735
      Left            =   2040
      TabIndex        =   7
      Top             =   4560
      Width           =   1815
   End
   Begin VB.CommandButton cmdpsych219 
      Caption         =   "Phychology219"
      Height          =   735
      Left            =   2160
      TabIndex        =   6
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton cmdns351 
      Caption         =   "NS351"
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   3360
      Width           =   1815
   End
   Begin VB.CommandButton cmdeducation111 
      Caption         =   "Education111"
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Width           =   1815
   End
   Begin VB.CommandButton cmdecon111 
      Caption         =   "Econ 111"
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1815
   End
   Begin VB.PictureBox picresults 
      Height          =   3615
      Left            =   4080
      ScaleHeight     =   3555
      ScaleWidth      =   3915
      TabIndex        =   2
      Top             =   360
      Width           =   3975
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "Clear"
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CommandButton cmdaccounting113 
      Caption         =   "accounting 113"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This program allows the user to select a number of particular classes
'The selected classes are shown in a picture box
Dim runningtotal As Single
Private Sub cmdaccounting113_Click()
Const accounting As Single = 3622.25
picresults.Print "Financial Accounting", FormatCurrency(accounting, 2)
runningtotal = runningtotal + accounting
End Sub
Private Sub cmdclear_Click()
'clear the picture box
picresults.Cls
runningtotal = 0
End Sub
Private Sub cmdecon111_Click()
Const economics As Single = 3622.25
picresults.Print "Intro to Economics", FormatCurrency(economics, 2)
runningtotal = runningtotal + economics
End Sub
Private Sub cmdeducation111_Click()
Const education As Single = 3675.25
picresults.Print "Teaching in a Diverse World", FormatCurrency(education, 2)
runningtotal = runningtotal + education
End Sub
Private Sub cmdmgmt201_Click()
Const management As Single = 3622.25
picresults.Print "Principles of Management", FormatCurrency(management, 2)
runningtotal = runningtotal + management
End Sub
Private Sub cmdns351_Click()
Const naturalscience As Single = 3650.25
picresults.Print "Intro to Nutrition", FormatCurrency(naturalscience, 2)
runningtotal = runningtotal + naturalscience
End Sub
Private Sub cmdpsych219_Click()
Const psychology As Single = 3622.25
picresults.Print "Political Psychology", FormatCurrency(psychology, 2)
runningtotal = runningtotal + psychology
End Sub
Private Sub cmdreturn_Click()
'Hide the registration form
'Written By John O'Grady
'Written 10-25-09
    frmmajorlist.Show
    frmForm1.Hide
End Sub
Private Sub cmdsoc201_Click()
Const sociology As Single = 3622.25
picresults.Print "Social Statistics", FormatCurrency(sociology, 2)
runningtotal = runningtotal + sociology
End Sub
Private Sub cmdspanish200_Click()
Const spanish As Single = 3722.25
picresults.Print "Intermediate Spanish", FormatCurrency(spanish, 2)
runningtotal = runningtotal + spanish
End Sub
Private Sub cmdtotal_Click()
picresults.Print "*******************************************"
picresults.Print "Subtotal", FormatCurrency(runningtotal, 2)
MsgBox ("Congratulations, You have successfully registered for Fall 2009!")
End Sub


