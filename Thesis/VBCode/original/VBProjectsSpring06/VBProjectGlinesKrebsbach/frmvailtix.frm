VERSION 5.00
Begin VB.Form frmvailtix 
   BackColor       =   &H00C00000&
   Caption         =   "Vail Lift Tickets"
   ClientHeight    =   9300
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14025
   ForeColor       =   &H00C00000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pivvail 
      Height          =   3735
      Left            =   2280
      Picture         =   "frmvailtix.frx":0000
      ScaleHeight     =   3675
      ScaleWidth      =   6555
      TabIndex        =   5
      Top             =   2880
      Width           =   6615
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back to Vail"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4920
      TabIndex        =   4
      Top             =   7680
      Width           =   1335
   End
   Begin VB.CommandButton cmdSenior 
      Caption         =   "Seniors(65+)"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8160
      TabIndex        =   3
      Top             =   1680
      Width           =   2175
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
      Height          =   735
      Left            =   4320
      TabIndex        =   2
      Top             =   1680
      Width           =   2295
   End
   Begin VB.CommandButton cmdChild 
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
      Height          =   735
      Left            =   480
      TabIndex        =   1
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label lblname 
      Caption         =   "By: Levi Glines and John Krebsbach"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   10680
      Width           =   2775
   End
   Begin VB.Label lblKids 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Kids under 4 ski for FREE!!!!"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      TabIndex        =   6
      Top             =   6840
      Width           =   8055
   End
   Begin VB.Label lblVail 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Vail Lift Tickets"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   7575
   End
End
Attribute VB_Name = "frmvailtix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim day, price As Integer
Dim A As Integer

Private Sub cmdAdult_Click()
    A = InputBox("How many days do you want to ski?(1-7)", "Input")
    Open App.Path & "\Vailadulttix.txt" For Input As #1
    Do Until EOF(1)
        Input #1, day, price
    Loop
    Close #1
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
    frmvailtix.Hide
    frmVail.Show
End Sub

Private Sub cmdChild_Click()
    A = InputBox("How many days do you want to ski?(1-7)", "Input")
    Open App.Path & "\Vailchildtix.txt" For Input As #1
    Do Until EOF(1)
        Input #1, day, price
    Loop
    Close #1
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

Private Sub cmdSenior_Click()
    A = InputBox("How many days do you want to ski?(1-7)", "Input")
    Open App.Path & "\VailSeniortix.txt" For Input As #1
    Do Until EOF(1)
        Input #1, day, price
    Loop
    Close #1
    If A = 1 Then
        MsgBox "Your ticket price total for one day is $71", , "One"
    ElseIf A = 2 Then
        MsgBox "your ticket price total for two days is $142", , "two"
    ElseIf A = 3 Then
        MsgBox "your ticket price total for three days is $213", , "three"
    ElseIf A = 4 Then
        MsgBox "your ticket price total for four days is $284", , "four"
         
    ElseIf A = 5 Then
        MsgBox "your ticket price total for five days is $355", , "five"
        
    ElseIf A = 6 Then
        MsgBox "your ticket price total for six days is $426", , "six"
        
    ElseIf A = 7 Then
        MsgBox "your ticket price total for seven days is $497", , "seven"
    End If
End Sub

