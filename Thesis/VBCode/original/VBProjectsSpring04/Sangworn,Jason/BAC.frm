VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00FF0000&
   ClientHeight    =   9435
   ClientLeft      =   2115
   ClientTop       =   570
   ClientWidth     =   9885
   LinkTopic       =   "Form4"
   ScaleHeight     =   9435
   ScaleWidth      =   9885
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Architect"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4920
      TabIndex        =   3
      Top             =   7560
      Width           =   1815
   End
   Begin VB.PictureBox picresults 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Architect"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   1200
      ScaleHeight     =   3675
      ScaleWidth      =   8475
      TabIndex        =   2
      Top             =   3240
      Width           =   8535
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Click Here to enter Weight"
      BeginProperty Font 
         Name            =   "Architect"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4560
      TabIndex        =   1
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   "Blood Alcohol Content"
      BeginProperty Font 
         Name            =   "Architect"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   900
      Left            =   2040
      TabIndex        =   0
      Top             =   480
      Width           =   6825
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : Blood Alcohol Content (Bloodalcohollevel.vbp)
'Form Name : Blood Alcohol Content (BAC.frm)
'Author : Jason Sangworn
'Date Written : March 15, 2004
'Purpose of Project : To calculate the blood alcohol content
                      'from the user by using the
                      'type of alcohol consumed
                      'number of drinks consumed
                      'time in which the drinks were consumed
                      'weight of the user
                      'and finally calculate the BAC of the user
                      
'Purpose of Form : 'have the user input the weight
                   'then calculate the BAC
Option Explicit

Dim W As Double
Dim BAC As Double

Private Sub Command1_Click()

W = InputBox("Please Enter Your Weight in Pounds", "Weight")
    
BAC = (Alcohol(A) * Number(J) * (0.075) / W) - (Time(T) * 0.015)

If BAC < 0 Then
    picresults.Print "Your Blood Alcohol Level is "; FormatNumber(BAC, 3)
    picresults.Print "There is a negligible amount of alcohol in your system.  You are not legally intoxicated."
ElseIf BAC > 0.1 Then
    picresults.Print "Your Blood Alcohol Level is "; FormatNumber(BAC, 3)
    picresults.Print "In All states you would be considered intoxicated and arrested for DUI."
ElseIf BAC >= 0.08 And BAC < 0.1 Then
    picresults.Print "Your Blood Alcohol Level is "; FormatNumber(BAC, 3)
    picresults.Print "In MOST states you would be considered intoxicated and arrested for DUI."
End If


Command1.Enabled = False

End Sub


Private Sub Command3_Click()
End
End Sub
