VERSION 5.00
Begin VB.Form frmNinth 
   Caption         =   "Your Standings"
   ClientHeight    =   8715
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10995
   LinkTopic       =   "Form1"
   ScaleHeight     =   8715
   ScaleWidth      =   10995
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   10095
      Left            =   0
      Picture         =   "Form9.frx":0000
      ScaleHeight     =   10035
      ScaleWidth      =   10995
      TabIndex        =   0
      Top             =   0
      Width           =   11055
      Begin VB.CommandButton cmdQuit 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Quit"
         BeginProperty Font 
            Name            =   "Chiller"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   8160
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   7200
         Width           =   1335
      End
      Begin VB.PictureBox picResults7 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Chiller"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   600
         ScaleHeight     =   555
         ScaleWidth      =   9795
         TabIndex        =   3
         Top             =   2760
         Width           =   9855
      End
      Begin VB.CommandButton cmdCalculate 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Show Me My Standings"
         BeginProperty Font 
            Name            =   "Chiller"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1680
         Width           =   4095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Your Standings"
         BeginProperty Font 
            Name            =   "Chiller"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1215
         Left            =   3360
         TabIndex        =   1
         Top             =   240
         Width           =   4815
      End
   End
End
Attribute VB_Name = "frmNinth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdCalculate_Click()
Dim Pos As Integer
For Pos = 1 To 4
    'print array 1, and 2,
    'add up points for array 2
    'Print out total of points
    Total = Total + TrickPoints1(Pos) + TrickPoints2(Pos) + TrickPoints3(Pos)
        
Next Pos
Select Case Total
    Case Is < 80
        picResults7.Print " Congratualtions " & PName & " You Received " & Total & " Points After 3 Runs Which Puts You In Last Place!"
    Case Is < 120
        picResults7.Print " Congratualtions " & PName & " You Received " & Total & " Points After 3 Runs Which Puts You In Fourth Place!"
    Case Is < 160
        picResults7.Print " Congratualtions " & PName & " You Received " & Total & " Points After 3 Runs Which Puts You In Third Place!"
    Case Is < 200
        picResults7.Print " Congratualtions " & PName & " You Received " & Total & " Points After 3 Runs Which Puts You In Second Place!"
    Case Is <= 240
        picResults7.Print " Congratualtions " & PName & " You Received " & Total & " Points After 3 Runs Which Puts You In First Place!"
End Select

End Sub

Private Sub cmdQuit_Click()
End
End Sub
