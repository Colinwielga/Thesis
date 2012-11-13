VERSION 5.00
Begin VB.Form frmhealth 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5985
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3165
   ScaleWidth      =   5985
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdback 
      Caption         =   "Back"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Width           =   1095
   End
   Begin VB.PictureBox picresults1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2040
      ScaleHeight     =   315
      ScaleWidth      =   2835
      TabIndex        =   6
      Top             =   1920
      Width           =   2895
   End
   Begin VB.PictureBox picresults 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   960
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   5
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton cmdcompute 
      Caption         =   "Click here to find out if you're healthy"
      Height          =   1095
      Left            =   3000
      TabIndex        =   4
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox txtrate 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   2040
      TabIndex        =   3
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox txtage 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   2040
      TabIndex        =   1
      Top             =   240
      Width           =   615
   End
   Begin VB.Label lbljeff 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Created By: Jeff Amble"
      Height          =   255
      Left            =   2760
      TabIndex        =   8
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Image imgheart 
      Height          =   1575
      Left            =   4200
      Picture         =   "frmhealth.frx":0000
      Top             =   120
      Width           =   1545
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Count your pulse for 10 seconds and enter the number here:"
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enter your Age:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmhealth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This form enables the user to enter their age and pulse and'
'receive their heart rate and whether or not they are healthy'
Option Explicit
'This button enables you to go back to the main page'
Private Sub cmdback_Click()
    frmhealth.Visible = False
    frmmain.Visible = True
End Sub
'This button computes the users heart rate'
Private Sub cmdcompute_Click()
    Dim A As Integer
    Dim R As Integer
    Dim B As Single
    A = Text1.Text
    R = Text2.Text
    B = (0.7 * (220 - A) + 0.3 * R)
    picresults.Print (0.7 * (220 - A) + 0.3 * R)
        Select Case B
            Case Is > 225
                picresults1.Print "You are Very Unhealthy"
            Case Is > 140
                picresults1.Print "You are Healthy"
            Case Is > 100
                picresults1.Print "You are Very Healthy"
            Case Is > 70
                picresults1.Print "You are Unhealthy"
            Case Is > 50
                picresults1.Print "You are Very Unhealthy"
            Case Else
                picresults1.Print "You may be dead"
        End Select
End Sub
'This picbox displays the users heart rate'
Private Sub picresults_Click()

End Sub
'This picbox displays the users health'
Private Sub picresults1_Click()

End Sub

