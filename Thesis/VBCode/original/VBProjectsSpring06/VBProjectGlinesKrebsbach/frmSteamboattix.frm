VERSION 5.00
Begin VB.Form frmSteamboattix 
   BackColor       =   &H00C00000&
   Caption         =   "Steamboat Lift Tickets"
   ClientHeight    =   9315
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13995
   ForeColor       =   &H00C00000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back to Steamboat"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   9840
      Width           =   1215
   End
   Begin VB.CommandButton cmdTeen 
      Caption         =   "Teen(13-17)"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5040
      TabIndex        =   5
      Top             =   6840
      Width           =   2295
   End
   Begin VB.PictureBox picpeak 
      Height          =   5655
      Left            =   2280
      Picture         =   "frmSteamboattix.frx":0000
      ScaleHeight     =   5595
      ScaleWidth      =   10155
      TabIndex        =   4
      Top             =   960
      Width           =   10215
   End
   Begin VB.CommandButton cmdSenior 
      Caption         =   "Seniors(65+)"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10440
      TabIndex        =   3
      Top             =   6840
      Width           =   2055
   End
   Begin VB.CommandButton cmdAdult 
      Caption         =   "Adult(18-64)"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7800
      TabIndex        =   2
      Top             =   6840
      Width           =   2055
   End
   Begin VB.CommandButton cmdChild 
      Caption         =   "Children(6-12)"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2280
      TabIndex        =   1
      Top             =   6840
      Width           =   2415
   End
   Begin VB.Label lblname 
      Caption         =   "By: Levi Glines and John Krebsbach"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   10680
      Width           =   2775
   End
   Begin VB.Label lblKids 
      BackStyle       =   0  'Transparent
      Caption         =   "Kids under 4 ski for FREE!!!!"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   7
      Top             =   7920
      Width           =   5175
   End
   Begin VB.Label lblTickets 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Steamboat Lift Tickets"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3960
      TabIndex        =   0
      Top             =   240
      Width           =   7575
   End
End
Attribute VB_Name = "frmSteamboattix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : Colorado Spring Break(Final.vbp)
'Form Name : frmKeystonetix(frmKeystonetix.frm)
'Author: Levi Glines and John Krebsbach
'Date : Thursday March 23, 2006
'Purpose of this form:  This form allows the user to look up the respective lift ticket
'prices for children, teens, adults, and seniors.the user can input the number of days they want
'to ski and they will get an output via a msgbox depending on which age they click on.
Option Explicit
Dim day, price As Integer
Dim A As Integer

Private Sub cmdAdult_Click()
    A = InputBox("How many days do you want to ski?(1-7)", "Input") 'sets the users input equal to A
    Open App.Path & "\Steamboatadulttix.txt" For Input As #1 'opens the text file for use
    Do Until EOF(1) ' allows for the entire file to be read
    Input #1, day, price 'inputs the data into arrays of day and price
    Loop 'reads the next line
    Close #1 'closes the file when finished load into arrays
    If A = 1 Then 'searches for the users input and displays price
        MsgBox "Your ticket price total for one day is $74", , "One"
    ElseIf A = 2 Then
        MsgBox "your ticket price total for two days is $148", , "two"
    ElseIf A = 3 Then
        MsgBox "your ticket price total for three days is $222", , "three"
    ElseIf A = 4 Then
        MsgBox "your ticket price total for four days is $296", , "four"
    ElseIf A = 5 Then
        MsgBox "your ticket price total for five days is $370", , "five"
    ElseIf A = 6 Then
        MsgBox "your ticket price total for six days is $414", , "six"
    ElseIf A = 7 Then
        MsgBox "your ticket price total for seven days is $483", , "seven"
    End If
End Sub

Private Sub cmdback_Click()
    frmSteamboattix.Hide 'hides this form
    frmSteamboat.Show 'brings you back to the steamboat form
End Sub

Private Sub cmdChild_Click()
    A = InputBox("How many days do you want to ski?(1-7)", "Input")
    Open App.Path & "\Steamboatchildtix.txt" For Input As #1
    Do Until EOF(1)
        Input #1, day, price
    Loop
    Close #1
    If A = 1 Then
        MsgBox "Your ticket price total for one day is $47", , "One"
    ElseIf A = 2 Then
        MsgBox "your ticket price total for two days is $94", , "two"
    ElseIf A = 3 Then
        MsgBox "your ticket price total for three days is $141", , "three"
    ElseIf A = 4 Then
        MsgBox "your ticket price total for four days is $188", , "four"
    ElseIf A = 5 Then
        MsgBox "your ticket price total for five days is $235", , "five"
    ElseIf A = 6 Then
        MsgBox "your ticket price total for six days is $277", , "six"
    ElseIf A = 7 Then
        MsgBox "your ticket price total for seven days is $319", , "seven"
    End If
End Sub

Private Sub cmdSenior_Click()
    A = InputBox("How many days do you want to ski?(1-7)", "Input")
    Open App.Path & "\Steamboatseniortix.txt" For Input As #1
    Do Until EOF(1)
        Input #1, day, price
    Loop
    Close #1
    If A = 1 Then
        MsgBox "Your ticket price total for one day is $56", , "One"
    ElseIf A = 2 Then
        MsgBox "your ticket price total for two days is $112", , "two"
    ElseIf A = 3 Then
        MsgBox "your ticket price total for three days is $168", , "three"
    ElseIf A = 4 Then
        MsgBox "your ticket price total for four days is $224", , "four"
    ElseIf A = 5 Then
        MsgBox "your ticket price total for five days is $280", , "five"
    ElseIf A = 6 Then
        MsgBox "your ticket price total for six days is $332", , "six"
    ElseIf A = 7 Then
        MsgBox "your ticket price total for seven days is $384", , "seven"
    End If
End Sub

Private Sub cmdTeen_Click()
    A = InputBox("How many days do you want to ski?(1-7)", "Input")
    Open App.Path & "\Steamboatteentix.txt" For Input As #1
    Do Until EOF(1)
    Input #1, day, price
    Loop
    Close #1
    If A = 1 Then
        MsgBox "Your ticket price total for one day is $56", , "One"
    ElseIf A = 2 Then
        MsgBox "your ticket price total for two days is $112", , "two"
    ElseIf A = 3 Then
        MsgBox "your ticket price total for three days is $168", , "three"
    ElseIf A = 4 Then
        MsgBox "your ticket price total for four days is $224", , "four"
         
    ElseIf A = 5 Then
        MsgBox "your ticket price total for five days is $280", , "five"
        
    ElseIf A = 6 Then
        MsgBox "your ticket price total for six days is $332", , "six"
        
    ElseIf A = 7 Then
        MsgBox "your ticket price total for seven days is $384", , "seven"
    End If
End Sub

