VERSION 5.00
Begin VB.Form Managers 
   BackColor       =   &H00C00000&
   Caption         =   "Managers, Coaches, Staff"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10965
   LinkTopic       =   "Form1"
   ScaleHeight     =   8310
   ScaleWidth      =   10965
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      Height          =   2175
      Left            =   360
      Picture         =   "Managers.frx":0000
      ScaleHeight     =   2115
      ScaleWidth      =   3195
      TabIndex        =   6
      Top             =   4680
      Width           =   3255
   End
   Begin VB.CommandButton cmdStaff 
      BackColor       =   &H0080FFFF&
      Caption         =   "Show Staff Members"
      BeginProperty Font 
         Name            =   "Eras Demi ITC"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3720
      Width           =   1935
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00C00000&
      BeginProperty Font 
         Name            =   "Century"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   2775
      Left            =   3960
      ScaleHeight     =   2715
      ScaleWidth      =   6795
      TabIndex        =   4
      Top             =   240
      Width           =   6855
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   4080
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "Managers.frx":32DF
      Top             =   3480
      Width           =   6735
   End
   Begin VB.PictureBox Picture1 
      Height          =   2055
      Left            =   2160
      Picture         =   "Managers.frx":36BA
      ScaleHeight     =   1995
      ScaleWidth      =   1395
      TabIndex        =   2
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Return To Homepage"
      BeginProperty Font 
         Name            =   "Eras Demi ITC"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7440
      Width           =   2775
   End
   Begin VB.CommandButton cmdShow 
      BackColor       =   &H0080FFFF&
      Caption         =   "Show Managers && Coaches"
      BeginProperty Font 
         Name            =   "Eras Demi ITC"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2640
      Width           =   1935
   End
   Begin VB.PictureBox Picture3 
      Height          =   3015
      Left            =   0
      Picture         =   "Managers.frx":446B
      ScaleHeight     =   2955
      ScaleWidth      =   3795
      TabIndex        =   7
      Top             =   0
      Width           =   3855
   End
End
Attribute VB_Name = "Managers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Project Name: Twins Baseball
' Form Name: Managers
' Authors: Jake Krisnik & Mike Foley
' Date Written: October 24, 2006
' Form Objective: To provide the user with the ability to view managers and coaches of
'                 the twins and also view other staff members including medical staff.
'                 This form also focuses on Ron Gardenhire the Manager of the Minnesota
'                 Twins. We have some pictures of him for the user to enjoy as well as
'                 a brief explanation of his career.
Option Explicit
Private Sub cmdReturn_Click()
' This command button allows the user to navigate away from the Managers form and return
' to the Homepage.
    HomePage.Show
    Managers.Hide
End Sub

Private Sub cmdShow_Click()
' This command button opens an array that reads in to the file information on Managers and
' coaches. It prints the results in an appending table and clears the results from previous
' input.
    Dim I As Single
    Dim CoachName(1 To 7) As String
    Dim CoachNumber(1 To 7) As Integer
    Dim CoachPosition(1 To 7) As String
    picResults.Cls
    Open App.Path & "\Manager.txt" For Input As #3
    I = 0
    Do While I < 7                          'Tells the program to fill the array with the appropriate information
        I = I + 1                           'This information is about the coaches for the Twins
        Input #3, CoachName(I), CoachNumber(I), CoachPosition(I)
        picResults.Print UCase(CoachName(I)); "          ", CoachNumber(I), UCase(CoachPosition(I))
    Loop
    Close #3
End Sub


Private Sub cmdStaff_Click()
' This command button opens an array that reads in to the file information on other staff
' members. It prints the results in an appending table and clears the results from previous
' input.
    Dim I As Single
    Dim StaffName(1 To 10) As String
    Dim StaffPosition(1 To 10) As String
    picResults.Cls
    Open App.Path & "\Staff.txt" For Input As #4
    I = 0
    Do While I < 10                     'Tells the program to go through the array and fill it'
        I = I + 1                       'this information is about the coaches and managers for the Twins
        Input #4, StaffName(I), StaffPosition(I)
        picResults.Print UCase(StaffName(I)); "          ", UCase(StaffPosition(I))
    Loop
    Close #4
End Sub
