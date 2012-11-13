VERSION 5.00
Begin VB.Form frmInfo 
   BackColor       =   &H00808000&
   Caption         =   "Form1"
   ClientHeight    =   6225
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9450
   LinkTopic       =   "Form1"
   ScaleHeight     =   6225
   ScaleWidth      =   9450
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRetrieve 
      Caption         =   "Get information"
      Height          =   735
      Left            =   7680
      TabIndex        =   5
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton cmdClick2 
      Caption         =   "Click to view in order of size"
      Height          =   1215
      Left            =   480
      TabIndex        =   4
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CommandButton cmdReadfile 
      Caption         =   "Click to view in order of creation"
      Height          =   1455
      Left            =   480
      TabIndex        =   3
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to main page"
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   5280
      Width           =   2175
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   7320
      TabIndex        =   1
      Top             =   5640
      Width           =   1335
   End
   Begin VB.PictureBox picResults 
      Height          =   4215
      Left            =   3120
      ScaleHeight     =   4155
      ScaleWidth      =   4275
      TabIndex        =   0
      Top             =   360
      Width           =   4335
   End
   Begin VB.Label lbl1 
      BackColor       =   &H00808000&
      Caption         =   "Click here to retrieve information"
      Height          =   615
      Left            =   7800
      TabIndex        =   6
      Top             =   1800
      Width           =   1455
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: quick Facts about the MIAC'
    'Form name:Info
    'Author:Alec Beatty'
    'Written 10/18/2009'
    'Objective: to give basic info about The MIAC'
    
Option Explicit
Dim School(1 To 30) As String
Dim Year(1 To 20000) As Single
Dim CTR As Integer
Dim Pass As Integer, Pos As Integer, Temp As Integer 'Need to Declare dimensions for use in more than one subroutine'
Dim J As Integer
Dim enrollment(1 To 15000) As Single



Private Sub cmdClick2_Click()
   Dim TempSchool As String, Tempenrollment As Single, TempYear As Single 'Need temp dimensions so you don't lose them in sorting'
   picResults.Cls
   
   For Pass = 1 To CTR - 1 '
    For Pos = 1 To CTR - Pass
        If enrollment(Pos) > enrollment(Pos + 1) Then
            Tempenrollment = enrollment(Pos)
            enrollment(Pos) = enrollment(Pos + 1)
            enrollment(Pos + 1) = Tempenrollment
            TempSchool = School(Pos)
            School(Pos) = School(Pos + 1)
            School(Pos + 1) = TempSchool
            TempYear = Year(Pos)
            Year(Pos) = Year(Pos + 1)
            Year(Pos + 1) = TempYear
        End If
       Next Pos
       Next Pass
       
        picResults.Print "School"; Tab(20); "Enrollment"; Tab(40); "Year founded"
        picResults.Print "-------------------------------------------------------------------------------------------------------"
       For J = 1 To CTR
        picResults.Print School(J); Tab(20); enrollment(J); Tab(40); Year(J)
        Next J
        
    
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdReadfile_Click()
     Dim TempSchool As String, TempYear As Single, Tempenrollment As Single
    picResults.Cls
   For Pass = 1 To CTR - 1
    For Pos = 1 To CTR - Pass
        If Year(Pos) > Year(Pos + 1) Then
            TempYear = Year(Pos)
            Year(Pos) = Year(Pos + 1)
            Year(Pos + 1) = TempYear
            TempSchool = School(Pos)
            School(Pos) = School(Pos + 1)
            School(Pos + 1) = TempSchool
            Tempenrollment = enrollment(Pos)
            enrollment(Pos) = enrollment(Pos + 1)
            enrollment(Pos + 1) = Tempenrollment
        End If
     Next Pos
    Next Pass
        picResults.Print "School"; Tab(20); "Year founded"; Tab(40); "Enrollment"
        picResults.Print "------------------------------------------------------------------------------------------------------------------------"
       For J = 1 To CTR
        picResults.Print School(J); Tab(20); Year(J); Tab(40); enrollment(J) 'print results
        Next J
End Sub

Private Sub cmdRetrieve_Click()
    CTR = 0
    
 Open App.Path & "\MIAC1.txt" For Input As #1
 
    Do Until EOF(1)
        CTR = CTR + 1
    Input #1, School(CTR), Year(CTR), enrollment(CTR)
        
   Loop
        MsgBox "Information Retrieved", , "Attention"
        Close #1 'Need to close the file or else it will keep reading it'
End Sub

Private Sub cmdReturn_Click()
    frmInfo.Hide
    frmMIAC.Show
End Sub
