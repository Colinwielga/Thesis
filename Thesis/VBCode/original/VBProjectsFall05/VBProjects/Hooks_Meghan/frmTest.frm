VERSION 5.00
Begin VB.Form frmTest 
   BackColor       =   &H00C0C000&
   Caption         =   "Self Report Survey"
   ClientHeight    =   7080
   ClientLeft      =   1740
   ClientTop       =   1305
   ClientWidth     =   9660
   LinkTopic       =   "Form1"
   ScaleHeight     =   7080
   ScaleWidth      =   9660
   Begin VB.TextBox txtWeight 
      Height          =   375
      Left            =   4800
      TabIndex        =   21
      Top             =   5520
      Width           =   1215
   End
   Begin VB.TextBox txtCry 
      Height          =   375
      Left            =   4800
      TabIndex        =   19
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox txtSelf 
      Height          =   375
      Left            =   4800
      TabIndex        =   17
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox txtFatigue 
      Height          =   375
      Left            =   4800
      TabIndex        =   15
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox txtFriends 
      Height          =   375
      Left            =   4800
      TabIndex        =   13
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txtInterest 
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   5520
      Width           =   1095
   End
   Begin VB.TextBox txtIrritable 
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   4320
      Width           =   1095
   End
   Begin VB.TextBox txtBlue 
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   5640
      TabIndex        =   5
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "Submit"
      Height          =   495
      Left            =   2520
      TabIndex        =   4
      Top             =   6480
      Width           =   1215
   End
   Begin VB.TextBox txtEat 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox txtSleep 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label lblWeight 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Have you noticed a change in your weight? Gain? Loss?"
      Height          =   495
      Left            =   4800
      TabIndex        =   20
      Top             =   4920
      Width           =   4095
   End
   Begin VB.Label lblCry 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Do you sometimes have crying episodes or times when you feel like crying?"
      Height          =   495
      Left            =   4800
      TabIndex        =   18
      Top             =   3720
      Width           =   4215
   End
   Begin VB.Label lblSelf 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Are you generally happy with yourself? Do you like you?"
      Height          =   495
      Left            =   4800
      TabIndex        =   16
      Top             =   2520
      Width           =   4215
   End
   Begin VB.Label lblFatigue 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Do you become easily fatigued?"
      Height          =   495
      Left            =   4800
      TabIndex        =   14
      Top             =   1320
      Width           =   4215
   End
   Begin VB.Label lblFriend 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Have you been withdrawing from friends, family and/or social activites?"
      Height          =   495
      Left            =   4800
      TabIndex        =   12
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label lblInterest 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Have you lost interest in things that you once enjoyed doing?"
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   4920
      Width           =   4095
   End
   Begin VB.Label lblIrritable 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Do you often feel irritable or annoyed?"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   3720
      Width           =   4095
   End
   Begin VB.Label lblBlues 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Have you felt 'down' or 'blue' for two weeks or more?"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   4095
   End
   Begin VB.Label lblEating 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Have you, or friends/family, noticed a chance in your eating habits? Eating more? Less?"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   4095
   End
   Begin VB.Label lblSleeping 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Do you have trouble falling asleep or staying asleep more than two nights a week?"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Test(frmTest)
'Created by Meghan Hooks
'10-30-05
    'This form holds the inventory questions. This form has the calculations
    'which form the result score.

Option Explicit
Private Sub cmdCancel_Click()
frmTest.Hide
frmProject.Show
End Sub

Private Sub cmdSubmit_Click()
Dim Sleep As String
Dim Eating As String
Dim Blue As String
Dim Irritable As String
Dim Interests As String
Dim Friends As String
Dim Fatigue As String
Dim Self As String
Dim Cry As String
Dim Weight As String

Sleep = txtSleep.Text
Eating = txtEat.Text
Blue = txtBlue.Text
Irritable = txtIrritable.Text
Interests = txtInterest.Text
Friends = txtFriends.Text
Fatigue = txtFatigue.Text
Self = txtSelf.Text
Cry = txtCry.Text
Weight = txtWeight.Text

If Sleep = "Yes" Then
    Results = Results + 1
ElseIf Sleep = "No" Then
    Results = Results + 0
End If
If Eating = "Yes" Then
    Results = Results + 1
ElseIf Eating = "No" Then
    Results = Results + 0
End If
If Blue = "Yes" Then
    Results = Results + 1
ElseIf Blue = "No" Then
    Results = Results + 0
End If
If Irritable = "Yes" Then
    Results = Results + 1
ElseIf Irritable = "No" Then
    Results = Results + 0
End If
If Interests = "Yes" Then
    Results = Results + 1
ElseIf Interests = "No" Then
    Results = Results + 0
End If
If Friends = "Yes" Then
    Results = Results + 1
ElseIf Friends = "No" Then
    Results = Results + 0
End If
If Fatigue = "Yes" Then
    Results = Results + 1
ElseIf Fatigue = "No" Then
    Results = Results + 0
End If
If Self = "No" Then
    Results = Results + 1
ElseIf Self = "Yes" Then
    Results = Results + 0
End If
If Cry = "Yes" Then
    Results = Results + 1
ElseIf Cry = "No" Then
    Results = Results + 0
End If
If Weight = "Yes" Then
    Results = Results + 1
ElseIf Weight = "No" Then
    Results = Results + 0
End If
frmTest.Hide
frmProject.Show
End Sub



Private Sub VScroll1_Change()
    frmTest.Top = -VScroll1.Value
    VScroll1.Top = 0
End Sub
