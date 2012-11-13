VERSION 5.00
Begin VB.Form frmBrecktix 
   BackColor       =   &H00C00000&
   Caption         =   "Breck Lift Tickets"
   ClientHeight    =   9375
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14475
   FillColor       =   &H00C00000&
   ForeColor       =   &H00C00000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   9375
   ScaleWidth      =   14475
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back to Breck"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9240
      TabIndex        =   6
      Top             =   3720
      Width           =   1335
   End
   Begin VB.PictureBox picBreck 
      Height          =   5295
      Left            =   5160
      Picture         =   "frmBrecktix.frx":0000
      ScaleHeight     =   5235
      ScaleWidth      =   2955
      TabIndex        =   4
      Top             =   1680
      Width           =   3015
   End
   Begin VB.CommandButton cmdSeniors 
      Caption         =   "Seniors (65+)"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   600
      TabIndex        =   3
      Top             =   5400
      Width           =   2535
   End
   Begin VB.CommandButton cmdAdult 
      Caption         =   "Adult(13-64)"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   600
      TabIndex        =   2
      Top             =   3480
      Width           =   2535
   End
   Begin VB.CommandButton cmdChildren 
      Caption         =   "Children(5-12)"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   600
      TabIndex        =   1
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Label lblname 
      Caption         =   "By: Levi Glines and John Krebsbach"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   10560
      Width           =   2775
   End
   Begin VB.Label lblChild 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Children 4 and under ski for free!!!"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   480
      TabIndex        =   5
      Top             =   7560
      Width           =   9975
   End
   Begin VB.Label lblBreck 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Breckenridge Lift Tickets"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   8775
   End
End
Attribute VB_Name = "frmBrecktix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : Colorado Spring Break(Final.vbp)
'Form Name : frmBrecktix(frmBrecktix.frm)
'Author: Levi Glines and John Krebsbach
'Date : Thursday March 23, 2006
'Purpose of this form:  This form allows the user to look up the respective lift ticket
'prices for children, adults, and seniors.the user can input the number of days they want
'to ski and they will get an output via a msgbox depending on which age they click on.
Option Explicit
Dim day, price As Integer
Dim A As Integer

Private Sub cmdAdult_Click()
    A = InputBox("How many days do you want to ski?(1-7)", "Input") 'sets the users input number equal to A
    Open App.Path & "\Breckadulttix.txt" For Input As #1 'opens the text file

    Do Until EOF(1)
        Input #1, day, price 'inputs the text file into arrays of day and price
    Loop 'goes through the file until it reaches the end
    Close #1 'closes the text file
    If A = 1 Then 'whatever the user inputs the respective price is displayed
        MsgBox "Your ticket price total for one day is $75", , "One"
    ElseIf A = 2 Then
        MsgBox "your ticket price total for two days is $150", , "two"
    ElseIf A = 3 Then
        MsgBox "your ticket price total for three days is $225", , "three"
    ElseIf A = 4 Then
        MsgBox "your ticket price total for four days is $300", , "four"
    ElseIf A = 5 Then
        MsgBox "your ticket price total for five days is $375", , "five"
    ElseIf A = 6 Then
        MsgBox "your ticket price total for six days is $450", , "six"
    ElseIf A = 7 Then
        MsgBox "your ticket price total for seven days is $525", , "seven"
    End If
End Sub

Private Sub cmdback_Click()
    frmBrecktix.Hide
    frmBreckenridge.Show
End Sub

Private Sub cmdChildren_Click()
    A = InputBox("How many days do you want to ski?(1-7)", "Input") 'sets the users input equal to A
    Open App.Path & "\Breckchildtix.txt" For Input As #1 'opens the text file for use
    Do Until EOF(1)
        Input #1, day, price 'inputs the info in the text file into arrays of day and price
    Loop 'goes through the data until the end of the file is reached
    Close #1 'closes the text file when done inputing
    If A = 1 Then 'the respective price is printed after searching through the respective days
        MsgBox "Your ticket price total for one day is $39", , "One"
    ElseIf A = 2 Then
        MsgBox "your ticket price total for two days is $78", , "two"
    ElseIf A = 3 Then
        MsgBox "your ticket price total for three days is $117", , "three"
    ElseIf A = 4 Then
        MsgBox "your ticket price total for four days is $156", , "four"
    ElseIf A = 5 Then
        MsgBox "your ticket price total for five days is $195", , "five"
    ElseIf A = 6 Then
        MsgBox "your ticket price total for six days is $234", , "six"
    ElseIf A = 7 Then
        MsgBox "your ticket price total for seven days is $273", , "seven"
    End If
End Sub

Private Sub cmdSeniors_Click()
    A = InputBox("How many days do you want to ski?(1-7)", "Input") 'sets the users input equal to A
    Open App.Path & "\Breckseniorstix.txt" For Input As #1 'opens the text file and sets it equal to #1
    Do Until EOF(1) 'lets the computer know to go until the end of the file
        Input #1, day, price 'loads the text file and puts it in arrays of days and price
    Loop
    Close #1
    If A = 1 Then 'the respective price is printed after searching through the respective days
        MsgBox "Your ticket price total for one day is $65", , "One"
    ElseIf A = 2 Then
        MsgBox "your ticket price total for two days is $130", , "two"
    ElseIf A = 3 Then
        MsgBox "your ticket price total for three days is $195", , "three"
    ElseIf A = 4 Then
        MsgBox "your ticket price total for four days is $260", , "four"
    ElseIf A = 5 Then
        MsgBox "your ticket price total for five days is $325", , "five"
    ElseIf A = 6 Then
        MsgBox "your ticket price total for six days is $390", , "six"
    ElseIf A = 7 Then
        MsgBox "your ticket price total for seven days is $455", , "seven"
    End If
End Sub
