VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   7395
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8715
   LinkTopic       =   "Form2"
   ScaleHeight     =   7395
   ScaleWidth      =   8715
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   4695
      Left            =   4440
      ScaleHeight     =   4635
      ScaleWidth      =   3075
      TabIndex        =   1
      Top             =   840
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   2055
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   2655
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Dim SwimmerTime(1 To 20) As Double
Dim SwimmerName(1 To 20) As String
Dim SwimmerEvent(1 To 20) As String
Dim SwimmerAge(1 To 20) As Integer


Dim Pass As Integer
Dim Pos As Integer
Dim Temp As Integer
Dim I As Integer
Dim CTR As Integer

Open App.Path & "/topteens50free.txt" For Input As #1

CTR = 0
Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, SwimmerName, SwimmerAge, SwimmerEvent, SwimmerTime
    Loop
Close #1
'For Pass = 1 To CTR - 1
    'For Pos = 1 To CTR - Pass
        'If SwimmerTime(Pos) > SwimmerTime(Pos + 1) Then
         '   Temp = SwimmerTime(Pos)
          '  SwimmerTime(Pos) = SwimmerTime(Pos + 1)
           ' SwimmerTime(Pos + 1) = Temp
            'End If
        'Next Pos
    'Next Pass
    
    Picture1.Print "Name", "Event", "Time"
    Picture1.Print "------------------------------------------------------------------------------------------------------------------"
For Pos = 1 To CTR
    Picture1.Print SwimmerName(Pos); Tab(20); , SwimmerEvent(Pos); Tab(20); , SwimmerTime(Pos); Tab(20);
Next Pos
End Sub
