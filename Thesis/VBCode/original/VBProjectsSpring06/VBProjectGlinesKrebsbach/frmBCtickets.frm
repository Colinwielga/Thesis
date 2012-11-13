VERSION 5.00
Begin VB.Form frmBCtickets 
   BackColor       =   &H00C00000&
   Caption         =   "Beaver Creek Lift Tickets"
   ClientHeight    =   8760
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14625
   FillColor       =   &H00FF0000&
   ForeColor       =   &H00C00000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8760
   ScaleWidth      =   14625
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back to Beaver Creek"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      TabIndex        =   5
      Top             =   8040
      Width           =   2775
   End
   Begin VB.PictureBox picBeaver 
      Height          =   4095
      Left            =   3480
      Picture         =   "frmBCtickets.frx":0000
      ScaleHeight     =   4035
      ScaleWidth      =   8235
      TabIndex        =   4
      Top             =   3480
      Width           =   8295
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
      Height          =   1335
      Left            =   10920
      TabIndex        =   3
      Top             =   1680
      Width           =   3135
   End
   Begin VB.CommandButton cmdAdults 
      Caption         =   "Adults(13 to 64)"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   6000
      TabIndex        =   2
      Top             =   1680
      Width           =   3255
   End
   Begin VB.CommandButton cmdChildren 
      Caption         =   "Children( 5 to 12)"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1080
      TabIndex        =   1
      Top             =   1680
      Width           =   3135
   End
   Begin VB.Label lblname 
      Caption         =   "By: Levi Glines and John Krebsbach"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   10680
      Width           =   2775
   End
   Begin VB.Label lblLift 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Beaver Creek Lift Tickets"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3600
      TabIndex        =   0
      Top             =   240
      Width           =   7695
   End
End
Attribute VB_Name = "frmBCtickets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : Colorado Spring Break(Final.vbp)
'Form Name : frmBCtickets(frmBCtickets.frm)
'Author: Levi Glines and John Krebsbach
'Date : Thursday March 23, 2006
'Purpose of this form:  This form allows the user to look up the respective lift ticket
'prices for children, adults, and seniors.the user can input the number of days they want
'to ski and they will get an output via a msgbox depending on which age they click on.
Option Explicit
Dim day, price As Integer
Dim A As Integer


Private Sub cmdAdults_Click()
    A = InputBox("How many days do you want to ski?(1-7)", "Input") 'sets the users input equal to A
    Open App.Path & "\BCadulttix.txt" For Input As #1 'opens the text file

    Do Until EOF(1)
        Input #1, day, price 'inputs the text file into an array of day and price
    Loop 'loops through the text file until the end of the file
    Close #1 'closes the text file
    If A = 1 Then
        MsgBox "Your ticket price total for one day is $81", , "One"
    ElseIf A = 2 Then
        MsgBox "your ticket price total for two days is $162", , "two"
    ElseIf A = 3 Then
        MsgBox "your ticket price total for three days is $243", , "three"
    ElseIf A = 4 Then
        MsgBox "your ticket price total for four days is $324", , "four"
         
    ElseIf A = 5 Then
        MsgBox "your ticket price total for five days is $405", , "five"
        
    ElseIf A = 6 Then
        MsgBox "your ticket price total for six days is $486", , "six"
    ElseIf A = 7 Then
        MsgBox "your ticket price total for seven days is $567", , "seven"
    End If
End Sub

Private Sub cmdback_Click()
    frmBCtickets.Hide 'hides this form
    frmBeaver.Show 'brings you back to Beaver form
End Sub
Private Sub cmdChildren_Click()
    A = InputBox("How many days do you want to ski?(1-7)", "Input") 'sets the users input equal to A
    Open App.Path & "\BCchildtix.txt" For Input As #1 'loads the text file
    Do Until EOF(1)
        Input #1, day, price 'inputs the text file into arrays of day and price
    Loop 'loops through the file until the end is reached
    Close #1 'closes the text file
    If A = 1 Then
        MsgBox "Your ticket price total for one day is $49", , "One"
    ElseIf A = 2 Then
        MsgBox "your ticket price total for two days is $98", , "two"
    ElseIf A = 3 Then
        MsgBox "your ticket price total for three days is $147", , "three"
    ElseIf A = 4 Then
        MsgBox "your ticket price total for four days is $196", , "four"
    ElseIf A = 5 Then
        MsgBox "your ticket price total for five days is $245", , "five"
    ElseIf A = 6 Then
        MsgBox "your ticket price total for six days is $294", , "six"
    ElseIf A = 7 Then
        MsgBox "your ticket price total for seven days is $343", , "seven"
    End If
End Sub
Private Sub cmdSeniors_Click()
    A = InputBox("How many days do you want to ski?(1-7)", "Input") 'sets the users input equal to A
    Open App.Path & "\BCseniorstix.txt" For Input As #1 'opens the text file and sets it equal to #1
    Do Until EOF(1)
        Input #1, day, price 'inputs the file into arrays of day and price
    Loop 'loops through the file until the end is reached
    Close #1 'closes the file
    If A = 1 Then ' if the users inputs a 1 $71 is printed
        MsgBox "Your ticket price total for one day is $71", , "One"
    ElseIf A = 2 Then ' if the user inputs a 2 $142 is printed
        MsgBox "your ticket price total for two days is $142", , "two"
    ElseIf A = 3 Then 'if the user inputs a 3: $213 is printed
        MsgBox "your ticket price total for three days is $213", , "three"
    ElseIf A = 4 Then 'if the user inputs a 4: $284 is printed
        MsgBox "your ticket price total for four days is $284", , "four"
    ElseIf A = 5 Then 'if the user inputs a 5: $355 is printed
        MsgBox "your ticket price total for five days is $355", , "five"
    ElseIf A = 6 Then 'if the user inputs a 6: $426 is printed
        MsgBox "your ticket price total for six days is $426", , "six"
    ElseIf A = 7 Then 'if the user inputs a 7: $497 is printed
        MsgBox "your ticket price total for seven days is $497", , "seven"
    End If
End Sub
