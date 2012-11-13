VERSION 5.00
Begin VB.Form frmWanted 
   BackColor       =   &H00FF8080&
   Caption         =   "Wanted"
   ClientHeight    =   7740
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8280
   LinkTopic       =   "Form1"
   ScaleHeight     =   7740
   ScaleWidth      =   8280
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search By Wanted Level"
      Height          =   735
      Left            =   6360
      TabIndex        =   11
      Top             =   4320
      Width           =   1575
   End
   Begin VB.CommandButton cmdLookUp 
      Caption         =   "Look Up More Info"
      Height          =   735
      Left            =   4440
      TabIndex        =   10
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton cmdWantedLevel 
      Caption         =   "By Wanted Level"
      Height          =   735
      Left            =   6360
      TabIndex        =   9
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton cmdAge 
      BackColor       =   &H8000000E&
      Caption         =   "By Age"
      Height          =   735
      Left            =   6360
      TabIndex        =   7
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton cmdName 
      Caption         =   "By Name"
      Height          =   735
      Left            =   6360
      TabIndex        =   6
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   735
      Left            =   840
      TabIndex        =   5
      Top             =   6840
      Width           =   1575
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Patrol Car"
      Height          =   735
      Left            =   6360
      TabIndex        =   4
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   735
      Left            =   3960
      TabIndex        =   3
      Top             =   6840
      Width           =   1575
   End
   Begin VB.PictureBox picOutput 
      Height          =   5055
      Left            =   360
      ScaleHeight     =   4995
      ScaleWidth      =   5595
      TabIndex        =   2
      Top             =   1560
      Width           =   5655
   End
   Begin VB.CommandButton cmdLoadStearns 
      Caption         =   "Load Stearn's Co"
      Height          =   735
      Left            =   2400
      TabIndex        =   1
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton cmdLoadFBI 
      Caption         =   "Load FBI"
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label lblSort 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "Sort By:"
      Height          =   255
      Left            =   6360
      TabIndex        =   8
      Top             =   720
      Width           =   1575
   End
End
Attribute VB_Name = "frmWanted"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Patrol Car
' frmWanted
' Kevin Conlon
' 11/4/08
' this form is to look up current wanted criminals
Dim LastName(1 To 100) As String
Dim FirstName(1 To 100) As String
Dim Age(1 To 100) As Integer
Dim Crime(1 To 100) As String
Dim Number(1 To 100) As Integer
Dim CTR As Integer

Private Sub cmdAge_Click()
    Dim Temp As Integer
    Dim TempLast As String
    Dim TempFirst As String
    Dim TempNumber As Integer
    Dim TempCrime As String
    Dim Pass As Integer
    Dim Pos As Integer
    Dim N As Integer
        
    For Pass = 1 To CTR - 1
        For Pos = 1 To CTR - Pass
            If Age(Pos) < Age(Pos + 1) Then
                Temp = Age(Pos)
                ' sort other arrays
                TempLast = LastName(Pos)
                TempFirst = FirstName(Pos)
                TempNumber = Number(Pos)
                TempCrime = Crime(Pos)
                Age(Pos) = Age(Pos + 1)
                ' sort other arrays
                LastName(Pos) = LastName(Pos + 1)
                FirstName(Pos) = FirstName(Pos + 1)
                Number(Pos) = Number(Pos + 1)
                Crime(Pos) = Crime(Pos + 1)
                Age(Pos + 1) = Temp
                ' sort other arrays
                LastName(Pos + 1) = TempLast
                FirstName(Pos + 1) = TempFirst
                Number(Pos + 1) = TempNumber
                Crime(Pos + 1) = TempCrime
            End If
        Next Pos
    Next Pass
    picOutput.Print "Order of data: Last Name, First Name, Age, Crime"
    picOutput.Print "*****************************************************************************"
    For N = 1 To CTR
        picOutput.Print LastName(N); ", "; FirstName(N); ", "; Age(N); ", "; Crime(N)
    Next N
End Sub

Private Sub cmdClear_Click()
    picOutput.Cls

End Sub

Private Sub cmdLoadFBI_Click()
    Open App.Path & "\FBITopTen.txt" For Input As #1
    CTR = 0
    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, LastName(CTR), FirstName(CTR), Age(CTR), Crime(CTR), Number(CTR)
    Loop
    Close #1
    
End Sub

Private Sub cmdLoadStearns_Click()
    Open App.Path & "\StearnsTopTen.txt" For Input As #2
    Do Until EOF(2)
        CTR = CTR + 1
        Input #2, LastName(CTR), FirstName(CTR), Age(CTR), Crime(CTR), Number(CTR)
    Loop
    Close #2
End Sub

Private Sub cmdLookUp_Click()
    NameLookUp = InputBox("Who do you want to know more about", "Name")
    Select Case NameLookUp
        Case "Bin Laden"
            frmPoints.Show
            frmWanted.Hide
        Case "Bulger"
            frmPoints.Show
            frmWanted.Hide
        Case Else
            MsgBox "Sorry, that name is not available.", , "Not Available"
    End Select
End Sub

Private Sub cmdName_Click()
    Dim TempAge As Integer
    Dim Temp As String
    Dim TempFirst As String
    Dim TempNumber As Integer
    Dim TempCrime As String
    Dim Pass As Integer
    Dim Pos As Integer
    Dim N As Integer
        
    For Pass = 1 To CTR - 1
        For Pos = 1 To CTR - Pass
            If LastName(Pos) > LastName(Pos + 1) Then
                Temp = LastName(Pos)
                ' sort other arrays
                TempAge = Age(Pos)
                TempFirst = FirstName(Pos)
                TempNumber = Number(Pos)
                TempCrime = Crime(Pos)
                LastName(Pos) = LastName(Pos + 1)
                ' sort other arrays
                Age(Pos) = Age(Pos + 1)
                FirstName(Pos) = FirstName(Pos + 1)
                Number(Pos) = Number(Pos + 1)
                Crime(Pos) = Crime(Pos + 1)
                LastName(Pos + 1) = Temp
                ' sort other arrays
                Age(Pos + 1) = TempAge
                FirstName(Pos + 1) = TempFirst
                Number(Pos + 1) = TempNumber
                Crime(Pos + 1) = TempCrime
            End If
        Next Pos
    Next Pass
    picOutput.Print "Order of data: Last Name, First Name, Age, Crime"
    picOutput.Print "*****************************************************************************"
    For N = 1 To CTR
        picOutput.Print LastName(N); ", "; FirstName(N); ", "; Age(N); ", "; Crime(N)
    Next N
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdReturn_Click()
    frmPatrolCar.Show
    frmWanted.Hide
End Sub

Private Sub cmdSearch_Click()
    Dim X As Integer
    X = InputBox("Please enter number (1 to 10)", "Number")
    MsgBox FirstName(X) & " " & LastName(X) & " is wanted level " & X, , "Wanted"
End Sub

Private Sub cmdWantedLevel_Click()
    Dim Temp As Integer
    Dim TempLast As String
    Dim TempFirst As String
    Dim TempAge As Integer
    Dim TempCrime As String
    Dim Pass As Integer
    Dim Pos As Integer
    Dim N As Integer
        
    For Pass = 1 To CTR - 1
        For Pos = 1 To CTR - Pass
            If Number(Pos) > Number(Pos + 1) Then
                Temp = Number(Pos)
                TempLast = LastName(Pos)
                TempFirst = FirstName(Pos)
                TempAge = Age(Pos)
                TempCrime = Crime(Pos)
                Number(Pos) = Number(Pos + 1)
                LastName(Pos) = LastName(Pos + 1)
                FirstName(Pos) = FirstName(Pos + 1)
                Age(Pos) = Age(Pos + 1)
                Crime(Pos) = Crime(Pos + 1)
                Number(Pos + 1) = Temp
                LastName(Pos + 1) = TempLast
                FirstName(Pos + 1) = TempFirst
                Age(Pos + 1) = TempAge
                Crime(Pos + 1) = TempCrime
            End If
        Next Pos
    Next Pass
    picOutput.Print "Order of data: Last Name, First Name, Age, Crime"
    picOutput.Print "***************************************************************************"
    For N = 1 To CTR
        picOutput.Print LastName(N); ", "; FirstName(N); ", "; Age(N); ", "; Crime(N)
    Next N
End Sub
