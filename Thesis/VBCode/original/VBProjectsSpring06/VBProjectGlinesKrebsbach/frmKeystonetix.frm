VERSION 5.00
Begin VB.Form frmKeystonetix 
   BackColor       =   &H00C00000&
   Caption         =   "Keystone Lift Tickets"
   ClientHeight    =   8910
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13920
   ForeColor       =   &H00C00000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back to Keystone "
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      TabIndex        =   6
      Top             =   8520
      Width           =   2775
   End
   Begin VB.PictureBox picsnow 
      Height          =   5415
      Left            =   2400
      Picture         =   "frmKeystonetix.frx":0000
      ScaleHeight     =   5355
      ScaleWidth      =   10635
      TabIndex        =   4
      Top             =   1920
      Width           =   10695
   End
   Begin VB.CommandButton cmdAdult 
      Caption         =   "Adult(13-65)"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      TabIndex        =   3
      Top             =   960
      Width           =   3135
   End
   Begin VB.CommandButton cmdSeniors 
      Caption         =   "Seniors(65+)"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9840
      TabIndex        =   2
      Top             =   960
      Width           =   3135
   End
   Begin VB.CommandButton cmdChildren 
      Caption         =   "Children(5-12)"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   960
      Width           =   3135
   End
   Begin VB.Label lblname 
      Caption         =   "By: Levi Glines and John Krebsbach"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   10680
      Width           =   2775
   End
   Begin VB.Label lblFree 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Children 4 and under ski for free!!!!"
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
      Left            =   2880
      TabIndex        =   5
      Top             =   7440
      Width           =   10095
   End
   Begin VB.Label lbltix 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Keystone Lift Tickets"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2760
      TabIndex        =   0
      Top             =   120
      Width           =   9855
   End
End
Attribute VB_Name = "frmKeystonetix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : Colorado Spring Break(Final.vbp)
'Form Name : frmKeystonetix(frmKeystonetix.frm)
'Author: Levi Glines and John Krebsbach
'Date : Thursday March 23, 2006
'Purpose of this form:  This form allows the user to look up the respective lift ticket
'prices for children, adults, and seniors.the user can input the number of days they want
'to ski and they will get an output via a msgbox depending on which age they click on.
Option Explicit
Dim day, price As Integer
Dim A As Integer

Private Sub cmdAdult_Click()
    A = InputBox("How many days do you want to ski?(1-7)", "Input") 'sets the user input equal to A
    Open App.Path & "\Keystoneadulttix.txt" For Input As #1 'opens the text file and sets it equal to #1
    Do Until EOF(1) 'allows for the file to be read from start to finish
        Input #1, day, price 'opens the file and loads data into arrays of day and price
    Loop 'goes to the next line in the file
    Close #1 'closes the file
    If A = 1 Then 'searches for the users input and displays the price
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
    frmKeystonetix.Hide 'hides this form
    frmKeystone.Show 'brings you back to the keystone form
End Sub

Private Sub cmdChildren_Click()
    A = InputBox("How many days do you want to ski?(1-7)", "Input") 'sets the users input equal to A
    Open App.Path & "\Keystonechildtix.txt" For Input As #1 'opens the text file for use
    Do Until EOF(1) 'allows for the whole file to be input
    Input #1, day, price 'loads the data into arrays of day and price
    Loop 'reads the next line in the file
    Close #1 'closes the file when finished
    If A = 1 Then 'searches for the users input and displays the respective price
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
    Open App.Path & "\Keystoneseniorstix.txt" For Input As #1 'opens the text file and sets it equal to #1
    Do Until EOF(1) 'allows for the entire file to be read
        Input #1, day, price 'inputs the data into arrays of day and price
    Loop 'reads the next line in the file
    Close #1 'closes the text file
    If A = 1 Then 'searches for the users input and displays the respective price
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
