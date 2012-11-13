VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10065
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   10065
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      Height          =   5775
      Left            =   3720
      ScaleHeight     =   5715
      ScaleWidth      =   5835
      TabIndex        =   4
      Top             =   600
      Width           =   5895
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   975
      Left            =   240
      TabIndex        =   3
      Top             =   3840
      Width           =   2295
   End
   Begin VB.CommandButton cmdSort 
      Caption         =   "Sort Data"
      Height          =   855
      Left            =   240
      TabIndex        =   2
      Top             =   2760
      Width           =   2295
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search for Snow Depth"
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   1680
      Width           =   2295
   End
   Begin VB.CommandButton cmdRead 
      Caption         =   "Read and Print Data"
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Snow(1 To 100) As Single
Dim Cities(1 To 100) As String
Dim CTR As Integer
Dim N As Integer
Dim Found As Boolean
Dim Depth As Single

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdRead_Click()
CTR = 0
    Open App.Path & "\snowdepth.txt" For Input As #1
    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, Snow(CTR), Cities(CTR)
        
    Loop
    Close #1

 picResults.Cls
    picResults.Print "Snow Depth", "Cities"
    picResults.Print "**************************************************************"
    
    For N = 1 To CTR
        picResults.Print Snow(N); Tab(15); Cities(N)
    Next N
End Sub

Private Sub cmdSearch_Click()
    Depth = InputBox("Enter Snow Depth", "Snow Depth")
    
    CTR = 0
    Found = False
    Do While Found = False And CTR < 100
        CTR = CTR + 1
        If Snow(CTR) < Depth Then
            Found = True
        End If
        
       
    Loop
    
    If Found Then
        MsgBox Cities(CTR) & " is the first city with less than that amount", , "Match"
    Else
        MsgBox "Sorry, no city has less than the amount you entered", , "No Match"
    End If
End Sub

Private Sub cmdSort_Click()
Dim Pass As Integer, Pos As Integer, Temp As Single, Temp2 As String
    picResults.Cls
    
    For Pass = 1 To CTR - 1
        For Pos = 1 To CTR - Pass
            If Snow(Pos) > Snow(Pos + 1) Then
                Temp = Snow(Pos)
                Snow(Pos) = Snow(Pos + 1)
                Snow(Pos + 1) = Temp
                    Temp2 = Cities(Pos)
                    Cities(Pos) = Cities(Pos + 1)
                    Cities(Pos + 1) = Temp2
                       
            End If
            
        Next Pos
    Next Pass
    
     picResults.Cls
    picResults.Print "Cities"; Tab(25); "Snow Depth"
    picResults.Print "**************************************************************"
    
    For N = 1 To CTR
        picResults.Print Cities(N); Tab(30); FormatNumber(Snow(N), 1)
    Next N
End Sub
