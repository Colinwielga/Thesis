VERSION 5.00
Begin VB.Form frmLucky 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Do You Feel Lucky?"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   Picture         =   "frmLucky.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdLastPage 
      Caption         =   "To Summary Page"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   11640
      TabIndex        =   9
      Top             =   4800
      Width           =   3375
   End
   Begin VB.CommandButton cmdHome 
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   13680
      TabIndex        =   8
      Top             =   7800
      Width           =   1335
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   13680
      TabIndex        =   7
      Top             =   9000
      Width           =   1335
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play!"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4920
      TabIndex        =   6
      ToolTipText     =   "Click me!"
      Top             =   3120
      Width           =   3615
   End
   Begin VB.TextBox txtNum5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   9960
      TabIndex        =   5
      Top             =   4080
      Width           =   615
   End
   Begin VB.TextBox txtNum4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8640
      TabIndex        =   4
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox txtNum3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   6480
      TabIndex        =   3
      Top             =   1440
      Width           =   615
   End
   Begin VB.TextBox txtNum2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   4080
      TabIndex        =   2
      Top             =   2280
      Width           =   615
   End
   Begin VB.TextBox txtNum1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2760
      TabIndex        =   1
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label lbl3 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Caption         =   "DO YOU FEEL LUCKY?"
      BeginProperty Font 
         Name            =   "Imprint MT Shadow"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   -360
      TabIndex        =   11
      Top             =   0
      Width           =   13935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "(Must be integers from 1 to 100)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   10
      Top             =   1200
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter your 5 favorite lottery numbers, then click ""Play!"""
      BeginProperty Font 
         Name            =   "Juice ITC"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   735
      Left            =   2400
      TabIndex        =   0
      Top             =   720
      Width           =   8295
   End
End
Attribute VB_Name = "frmLucky"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: The Game Of Life
'Form Name: frmLucky
'Authors: Pam Arndt and Alisa Davis
'Date Written: 3/4/09
'Objective: User enters 5 lucky numbers that go through a series of somewhat random mathematical formulas to determine different
'monetary or other prizes, which are saved for the summary page
Option Explicit

Private Sub cmdHome_Click()
'Back to Finer things in Life form
frmTheFinerThingsInLife.Show
frmLucky.Hide
End Sub

Private Sub cmdLastPage_Click()
frmSummary.Show
frmLucky.Hide 'takes user to summary page

End Sub

Private Sub cmdPlay_Click()
'This button will take the users 5 numbers and randomly award them money, prizes, vacations, or nothing! Take a chance!
Dim Num1 As Integer, Num2 As Integer, Num3 As Integer, Num4 As Integer, Num5 As Integer, Prize As Integer

Num1 = txtNum1.Text
Num2 = txtNum2.Text
Num3 = txtNum3.Text
Num4 = txtNum4.Text
Num5 = txtNum5.Text
Randomize

If (Num1 >= 1 And Num1 <= 100) And (Num2 >= 1 And Num2 <= 100) And (Num3 >= 1 And Num3 <= 100) And (Num4 >= 1 And Num4 <= 100) And (Num5 >= 1 And Num5 <= 100) Then
    If Num3 > Rnd * 10 * Num4 Then
        MsgBox "Congratulations! You won a free Caribbean Cruise!", , "Prize"
        Lottery = "Caribbean Cruise Trip"
    ElseIf Num2 <= Sqr(Rnd(50)) Then
        MsgBox "Congratulations! You won an African Safari Vacation!", , "Prize"
        Lottery = "African Safari Vacation"
    ElseIf Num1 = 13 Or Num2 = 13 Or Num3 = 13 Or Num4 = 13 Or Num5 = 13 Then
        MsgBox "Sorry, you must not be lucky.  No prize awarded!", , "Too bad"
        Lottery = "$0"
    ElseIf Num1 < (Rnd + Rnd * Num5) Then
        MsgBox "Congratulations! You won a Mustang Convertable!", , "Prize"
        Lottery = "Mustang Convertable"
    Else
        Prize = Abs((Int(Sqr(Rnd * Num1 * Num2 * Num3 * Num4 * Num5)) / Sqr(Rnd)) + Rnd)
        MsgBox "Congratulations! You won " & FormatCurrency(Prize), , "Prize"
        Lottery = FormatCurrency(Prize)
    End If
Else: MsgBox "Numbers must be integers from 1 to 100", , "Error"
End If

End Sub

Private Sub cmdQuit_Click()
End
End Sub
