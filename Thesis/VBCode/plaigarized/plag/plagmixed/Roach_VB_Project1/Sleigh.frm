VERSION 5.00
Begin VB.Form Sleigh 
   BackColor       =   &H80000009&
   Caption         =   "Form1"
   ClientHeight    =   13275
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17010
   LinkTopic       =   "Form1"
   ScaleHeight     =   13275
   ScaleWidth      =   17010
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSpeed 
      Caption         =   "By Speed"
      Height          =   975
      Left            =   12600
      TabIndex        =   9
      Top             =   11160
      Width           =   3375
   End
   Begin VB.CommandButton cmdIntel 
      Caption         =   "By Intelligence"
      Height          =   975
      Left            =   8520
      TabIndex        =   8
      Top             =   11160
      Width           =   3375
   End
   Begin VB.CommandButton cmdSTR 
      Caption         =   "By Strength"
      Height          =   975
      Left            =   12600
      TabIndex        =   7
      Top             =   9840
      Width           =   3375
   End
   Begin VB.CommandButton cmdLB 
      Caption         =   "By Weight"
      Height          =   975
      Left            =   8520
      TabIndex        =   6
      Top             =   9840
      Width           =   3375
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00000080&
      Caption         =   "Quit"
      Height          =   1095
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   11160
      Width           =   3135
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00008000&
      Caption         =   "Go Back"
      Height          =   1095
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   11160
      Width           =   3135
   End
   Begin VB.CommandButton cmdABC 
      Caption         =   "Alphabetical"
      Height          =   975
      Left            =   12600
      TabIndex        =   3
      Top             =   8520
      Width           =   3375
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load Reindeer"
      Height          =   975
      Left            =   8520
      TabIndex        =   2
      Top             =   8520
      Width           =   3375
   End
   Begin VB.PictureBox picResults 
      Height          =   6255
      Left            =   10560
      ScaleHeight     =   6195
      ScaleWidth      =   5475
      TabIndex        =   0
      Top             =   1800
      Width           =   5535
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Caption         =   "Help Santa!!! He Needs to Know What Reindeer Go Where?!"
      BeginProperty Font 
         Name            =   "Imprint MT Shadow"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   975
      Left            =   480
      TabIndex        =   1
      Top             =   240
      Width           =   15015
   End
   Begin VB.Image Image1 
      Height          =   6210
      Left            =   360
      Picture         =   "Sleigh.frx":0000
      Top             =   1680
      Width           =   9300
   End
End
Attribute VB_Name = "Sleigh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdABC_Click()
Dim PassA As Long
Dim CA As Long
Dim TempA As String

For PassA = 1 To 8
    For CA = 1 To 9 - PassA
            If Deer(CA) > Deer(CA + 1) Then
                    Temp = Deer(CA)






                    Deer(CA) = Deer(CA + 1)
                    Deer(CA + 1) = Temp
                    
            End If
    Next CA
Next PassA

    picResults.Cls
    picResults.Print
    picResults.Print "Name"






    picResults.Print "**************"

For CA = 1 To 9
    picResults.Print Deer(CA)
Next CA
End Sub

Private Sub cmdBack_Click()
    MainMenu.Show
    Sleigh.Hide
End Sub

Private Sub cmdIntel_Click()
Dim PassC As Long = 1









Dim CC As Long
Dim TempC As String

Do
    For CC = 1 To 9 - PassC
            If Intelligence(CC) < Intelligence(CC + 1) Then
                    TempC = Intelligence(CC)









                    Intelligence(CC) = Intelligence(CC + 1)
                    Intelligence(CC + 1) = TempC
                    
                    TempC = Deer(CC)
                    Deer(CC) = Deer(CC + 1)
                    Deer(CC + 1) = TempC
                    
            End If
    Next CC
Until PassC > 8

    picResults.Cls
    picResults.Print
    picResults.Print "Name"; Tab(11); "IQ"








    picResults.Print "************************"

For CC = 1 To 9
    picResults.Print Deer(CC); Tab(10); Intelligence(CC)
Next CC
End Sub

Private Sub cmdLB_Click()
Dim Pass, C As Integer
Dim Temp As String

Pass = 1
Do While Pass <= 8
    For C = 1 To 9 - Pass
            If Weight(C) < Weight(C + 1) Then








                    Temp = Weight(C)
                    Weight(C) = Weight(C + 1)
                    Weight(C + 1) = Temp
                    
                    Temp = Deer(C)
                    Deer(C) = Deer(C + 1)
                    Deer(C + 1) = Temp
                    
            End If
    Next C
Loop

    picResults.Cls
    picResults.Print
    picResults.Print "Name"; Tab(10); "Weight"
    picResults.Print "************************"





For C = 1 To 9
    picResults.Print Deer(C); Tab(10); Weight(C)
Next C
End Sub

Private Sub cmdLoad_Click()
    Dim I As Integer
    
    
    Open App.Path & "\reindeer.txt" For Input As #1
    
    picResults.Print









    picResults.Print "Name"; Tab(10); "Weight"; Tab(21); "IQ"; Tab(29); "Strength"; Tab(40); "Speed"
    picResults.Print "******************************************************************************"
    

    For I = 1 To 9
' telskjf asfkd slfsk
' asldkfj dlfksd f
' asldkf jdslfisdfd
        Input #1, Deer(I), Weight(I), Intelligence(I), Strength(I), Speed(I)
        picResults.Print Deer(I); Tab(10); Weight(I); Tab(20); Intelligence(I); Tab(30); Strength(I); Tab(40); Speed(I)
    Next I
    
    cmdABC.Enabled = True
    cmdSTR.Enabled = True









    cmdLB.Enabled = True
    cmdIntel.Enabled = True
    cmdSpeed.Enabled = True
' telskjf asfkd slfsk
' asldkfj dlfksd f
' asldkf jdslfisdfd
    cmdLoad.Enabled = False
End Sub


Private Sub cmdQuit_Click()
'Project Name: Rudolph
'Form Name: Sleigh
'Author: Patrick Roach











'Date written: February 25, 2010
'Objective: List at the reindeer by name, weight, strength, intelligence (IQ), and speed. Allows the user to sort the information by choice
    End
End Sub

Private Sub cmdSpeed_Click()
' telskjf asfkd slfsk
' asldkfj dlfksd f
' asldkf jdslfisdfd
Dim PassD, CD As Integer
Dim TempD As String

For PassD = 1 To 8
    For CD = 1 To 9 - PassD
            If Speed(CD) < Speed(CD + 1) Then







                    TempD = Speed(CD)
                    Speed(CD) = Speed(CD + 1)
                    Speed(CD + 1) = TempD






                    
                    TempD = Deer(CD)
                    Deer(CD) = Deer(CD + 1)
                    Deer(CD + 1) = TempD
                    
            End If
    Next CD
Next PassD

    picResults.Cls
    picResults.Print
' telskjf asfkd slfsk
' asldkfj dlfksd f
' asldkf jdslfisdfd
    picResults.Print "Name"; Tab(10); "Speed"
    picResults.Print "************************"

For CD = 1 To 9
    picResults.Print Deer(CD); Tab(10); Speed(CD)
Next CD
End Sub

Private Sub cmdSTR_Click()
Dim PassB, CB As Integer
Dim TempB As String

For PassB = 1 To 8
    For CB = 1 To 9 - PassB
            If Strength(CB) < Strength(CB + 1) Then
                    TempB = Strength(CB)




' telskjf asfkd slfsk
' asldkfj dlfksd f
' asldkf jdslfisdfd




                    Strength(CB) = Strength(CB + 1)
                    Strength(CB + 1) = TempB
                    
                    TempB = Deer(CB)
                    Deer(CB) = Deer(CB + 1)
                    Deer(CB + 1) = TempB
' telskjf asfkd slfsk
' asldkfj dlfksd f
' asldkf jdslfisdfd
                    
            End If
    Next CB
Next PassB

    picResults.Cls







    picResults.Print
    picResults.Print "Name"; Tab(10); "Strength"
    picResults.Print "************************"

For CB = 1 To 9
    picResults.Print Deer(CB); Tab(10); Strength(CB)
Next CB
End Sub
