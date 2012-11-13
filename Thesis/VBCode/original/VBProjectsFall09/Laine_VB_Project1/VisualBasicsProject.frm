VERSION 5.00
Begin VB.Form frmStage2 
   BackColor       =   &H80000012&
   Caption         =   "Stage2"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10905
   ForeColor       =   &H80000017&
   LinkTopic       =   "Form2"
   Picture         =   "VisualBasicsProject.frx":0000
   ScaleHeight     =   9495
   ScaleWidth      =   10905
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRank 
      BackColor       =   &H00FF0000&
      Caption         =   "Get Rank"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   4590
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2025
      Width           =   2580
   End
   Begin VB.TextBox txtFeet 
      Height          =   555
      Left            =   2565
      TabIndex        =   3
      Top             =   1485
      Width           =   2310
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FF0000&
      Caption         =   "The End"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   7965
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7560
      Width           =   1905
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H80000013&
      Height          =   4590
      Left            =   6210
      ScaleHeight     =   4530
      ScaleWidth      =   4395
      TabIndex        =   1
      Top             =   2835
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      Caption         =   "How do you rank in the MIAC? "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3240
      TabIndex        =   4
      Top             =   0
      Width           =   4065
   End
   Begin VB.Label lblHeight 
      BackColor       =   &H80000012&
      Caption         =   "Enter height In Feet"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   540
      Left            =   135
      TabIndex        =   0
      Top             =   945
      Width           =   3120
   End
End
Attribute VB_Name = "frmStage2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ctr2 As Integer


Private Sub cmdRank_Click()
 Dim feet As Single
 Dim n As Integer
    'get feet from textbox and assign to variable
    ctr2 = 1
    feet = Left(txtFeet.Text, 2)
    
    
    'How the Vaulter did'
    If feet >= 15 Then
            picResults.Print "You are the Best Pole Vaulter in the M.I.A.C"
        ElseIf feet >= 14.5 Then
            picResults.Print "You are among the top 5 Pole Vaulters in the M.I.A.C."
        ElseIf feet >= 14 Then
            picResults.Print "You are Doing great but need to improve!"
        ElseIf feet >= 13.6 Then
            picResults.Print "You've got a lot of work to do to become the best!"
        ElseIf feet >= 13 Then
            picResults.Print "You cleared 2 hieghts, but weren't very good!"
        ElseIf feet >= 12.6 Then
            picResults.Print "You cleared a height but need a large amount of work in order to become great"
        ElseIf feet <= 12 Then
            picResults.Print "You had a terrible meet"
        Else: picResults.Print "Error in hieght."
              picResults.Print "Invalid height entered!"
    End If
    
    picResults.Print
    picResults.Print feet; "was your height in FT"
    txtFeet.Text = " "
    
    feet = feet / 3.2808399
        Do While feet < metersArray(ctr2)
        ctr2 = ctr2 + 1
        Loop
        picResults.Print
        picResults.Print "You would have taken"; ctr2; "Place had you vaulted in the MIAC"
        
        
End Sub

Private Sub cmdQuit_Click()
End
End Sub



