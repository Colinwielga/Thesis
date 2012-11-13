VERSION 5.00
Begin VB.Form frmBar 
   BackColor       =   &H0080C0FF&
   Caption         =   "Form1"
   ClientHeight    =   9570
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   12105
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   ScaleHeight     =   9570
   ScaleWidth      =   12105
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtBox 
      Height          =   1335
      Left            =   3240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   8160
      Width           =   5655
   End
   Begin VB.CommandButton cmdFacts 
      BackColor       =   &H008080FF&
      Caption         =   "Fun Facts"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6120
      Width           =   1575
   End
   Begin VB.CommandButton cmdTrivia 
      BackColor       =   &H00C000C0&
      Caption         =   "Trivia!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5640
      Width           =   1575
   End
   Begin VB.CommandButton CmdQuit 
      BackColor       =   &H0000FFFF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8280
      Width           =   1335
   End
   Begin VB.CommandButton CmdGoBack 
      BackColor       =   &H00FFFF00&
      Caption         =   "Go Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1140
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8160
      Width           =   1290
   End
   Begin VB.Image imgBurger 
      Height          =   5580
      Left            =   1920
      Picture         =   "frmBar.frx":0000
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   8415
   End
   Begin VB.Label lblBar 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      Caption         =   "Mallie's Sports Bar and Grill"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2160
      TabIndex        =   2
      Top             =   1680
      Width           =   7815
   End
   Begin VB.Label lblBigbruger 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   """Home of the World's Biggest Burger"""
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2160
      TabIndex        =   1
      Top             =   960
      Width           =   7815
   End
   Begin VB.Label lblDetroit 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Detroit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2160
      TabIndex        =   0
      Top             =   240
      Width           =   7815
   End
End
Attribute VB_Name = "frmBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim J As Integer
'Man vs. Food
'frmBar
'Ty Nimens and Josh Seaburg
'February 2010
'Have information on the episode and make a button that displays information and has a trivia question that comes up in an inputbox


Private Sub cmdFacts_Click()
Dim Facts(1 To 20) As String, CTR As Integer
'fun facts was inspired by Fun Facts from Minnesota Tourism vbproject

Open App.Path & "\FunFacts.txt" For Input As #1
    Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, Facts(CTR)
Loop
Close #1
J = J + 1
If J = 9 Then
    J = 1
End If


'this code sample illustrates the use of a
'vertical scroll bar on a text box

'The text box properties that needed to be changed:
' set multiline to true
' set scrollbars to "2 - Vertical"

    txtBox.Text = txtBox.Text & vbCrLf & Facts(J)

End Sub

Private Sub cmdGoback_Click()
    frmBar.Hide
    frmMap.Show
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdTrivia_Click()
    Dim Guess As Single, X As Single
' This code relays to the user whether they are right or wrong on a trivia question, if wrong aproximatly how far off

    Guess = InputBox("How heavy is the world's biggest burger?", "Question?")
    
    Select Case Guess
        Case 185
            MsgBox ("Awesome guess, right on the money!")
        Case 175 To 195
            MsgBox ("You were less than or equal to 10 pounds off.")
        Case 165 To 205
            MsgBox ("You were less than or equal to 20 pounds off.")
        Case 135 To 235
            MsgBox ("You were less than or equal to 50 pounds off.")
        Case Else
            MsgBox ("You're out of the ball park!")
    End Select
    
End Sub
'textbox with scrollbar, make a notepad for fun facts and you click each time it reads it in the textbox then scroll down
Private Sub txtBox_Change()

End Sub
