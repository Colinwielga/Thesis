VERSION 5.00
Begin VB.Form frmBACalc 
   BackColor       =   &H00008000&
   Caption         =   "Form1"
   ClientHeight    =   8895
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   11130
   LinkTopic       =   "Form1"
   ScaleHeight     =   8895
   ScaleWidth      =   11130
   Begin VB.CommandButton cmdCalc 
      BackColor       =   &H0080C0FF&
      Caption         =   "Calculate BAC"
      Enabled         =   0   'False
      Height          =   1695
      Left            =   5880
      TabIndex        =   11
      Top             =   5400
      Width           =   3255
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   4560
      ScaleHeight     =   3315
      ScaleWidth      =   5955
      TabIndex        =   10
      Top             =   1800
      Width           =   6015
   End
   Begin VB.TextBox txtHours 
      BackColor       =   &H0080C0FF&
      Height          =   495
      Left            =   3240
      TabIndex        =   7
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H0080C0FF&
      Caption         =   "Female"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      TabIndex        =   6
      Top             =   5640
      Width           =   1335
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Male"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      TabIndex        =   5
      Top             =   5040
      Width           =   1095
   End
   Begin VB.TextBox txtWeight 
      BackColor       =   &H0080C0FF&
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   3600
      Width           =   1695
   End
   Begin VB.TextBox txtNumber 
      BackColor       =   &H0080C0FF&
      Height          =   495
      Left            =   2640
      TabIndex        =   1
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackColor       =   &H0080C0FF&
      Caption         =   "Please Select Gender"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   9
      Top             =   4560
      Width           =   2775
   End
   Begin VB.Label Label4 
      BackColor       =   &H0080C0FF&
      Caption         =   "Over How Many Hours?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   8
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080C0FF&
      Caption         =   "Enter Weight (Pounds)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   4
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080C0FF&
      Caption         =   "Enter Number of Drinks Consumed (12 oz.)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   2
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Blood Alcohol Content Calculator"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   8535
   End
   Begin VB.Menu file 
      Caption         =   "File"
      Begin VB.Menu Return 
         Caption         =   "Return to Menu"
      End
      Begin VB.Menu quit 
         Caption         =   "Quit"
      End
   End
End
Attribute VB_Name = "frmBACalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Gender As Single
'Crime Awareness Project
'frm License
'By: Alex Warzecha & Andrew Enzler
'Written: 3/18/2009
'The objective of this form is to offer a user friendly manner of calculating BAC.

Private Sub Check1_Click()
'If box is checked it is denoted by the value being set to 1 hence gender = 1
If Check1.Value = "1" Then
    Gender = 1
    cmdCalc.Enabled = True
ElseIf Check1.Value = "2" Then
cmdCalc.Enabled = True
End If
End Sub


Private Sub Check2_Click()
'If box is checked it is denoted by the value being set to 1 hence gender = 1
If Check2.Value = "1" Then
    Gender = 1.27
    cmdCalc.Enabled = True
ElseIf Check2.Value = "2" Then
cmdCalc.Enabled = True
End If
End Sub

Private Sub cmdCalc_Click()
'Denote variables
Dim Weight As Single, Number As Single, Hours As Single, BAC As Single
Dim secret As Single
Weight = txtWeight
Number = txtNumber
Hours = txtHours
'Compare potential weights and assign the given number as it relates to the increse
'in BAC per alcoholic drink.

Select Case Weight
    Case Is <= 50
        secret = 0.05
    Case Is <= 100
        secret = 0.04
    Case Is <= 120
        secret = 0.034
    Case Is <= 140
        secret = 0.021
    Case Is <= 160
        secret = 0.025
    Case Is <= 180
        secret = 0.023
    Case Is <= 200
        secret = 0.02
    Case Is <= 220
        secret = 0.018
    Case Is <= 240
        secret = 0.015
    Case Is <= 280
        secret = 0.012
    Case Is <= 310
        secret = 0.01
    Case Is <= 350
        secret = 0.007
    Case Else
        picResults.Print "Incorrect Weight Input."
End Select

'spells out the formula for BAC

BAC = ((Number * secret) - (0.016 * Hours)) * Gender

'prints results of the formula and offers appropriate judgement for driving safety.

picResults.Print "Your BAC is "; FormatNumber(BAC, 2)
picResults.Print " "
If BAC >= 0.08 Then
    picResults.Print "A person with a Blood Alcohol Content of "; FormatNumber(BAC, 2); " should not be driving."
picResults.Print " "
ElseIf BAC < 0.08 Then
    picResults.Print "A person with Blood Alcohol Content of "; FormatNumber(BAC, 2); " is legally allowed to drive."
picResults.Print " "
End If

End Sub

'offers option to quit

Private Sub quit_Click()
    End
End Sub

'Offers option for drop down file option.

Private Sub return_Click()
frmBACalc.Hide
frmHome.Show
End Sub
