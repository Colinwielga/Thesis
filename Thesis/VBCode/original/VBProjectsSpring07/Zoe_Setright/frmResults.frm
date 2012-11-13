VERSION 5.00
Begin VB.Form frmResults 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Form1"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9630
   FillColor       =   &H00FFFF00&
   ForeColor       =   &H00FFFF00&
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   9630
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSortName 
      BackColor       =   &H00FFFF80&
      Caption         =   "Sort by Name"
      Height          =   1455
      Left            =   7440
      TabIndex        =   3
      Top             =   2160
      Width           =   1935
   End
   Begin VB.CommandButton cmdSortYear 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Sort by Year"
      Height          =   1215
      Left            =   7440
      MaskColor       =   &H00C0C000&
      TabIndex        =   2
      Top             =   480
      Width           =   1935
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FFFF80&
      Caption         =   "Quit"
      Height          =   1935
      Left            =   7440
      TabIndex        =   1
      Top             =   8760
      Width           =   1815
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFFC0&
      Height          =   10575
      Left            =   360
      ScaleHeight     =   10515
      ScaleWidth      =   6795
      TabIndex        =   0
      Top             =   120
      Width           =   6855
   End
End
Attribute VB_Name = "frmResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdSortName_Click()
     'This button sorts the members by their name.
    Dim WTO(1 To 1000) As String
    Dim Month(1 To 1000) As String
    Dim Year(1 To 1000) As Single
    Dim TempYear As Single
    Dim TempWTO As String
    Dim TempMonth As String
    Dim Ctr As Integer
    Dim Pos As Integer
    Dim Pass As Integer
    Dim I As Integer
    
    Open App.Path & "\wto.txt" For Input As #1
    
    Ctr = 0
    Do Until EOF(1)
        Ctr = Ctr + 1
        Input #1, WTO(Ctr), Month(Ctr), Year(Ctr)
    Loop
    Close #1
    picResults.Cls
    For Pass = 1 To Ctr - 1
        For Pos = 1 To Ctr - Pass
            If WTO(Pos) > WTO(Pos + 1) Then
                TempWTO = WTO(Pos)
                WTO(Pos) = WTO(Pos + 1)
                WTO(Pos + 1) = TempWTO
                
                TempYear = Year(Pos)
                Year(Pos) = Year(Pos + 1)
                Year(Pos + 1) = TempYear
                
                TempMonth = Month(Pos)
                Month(Pos) = Month(Pos + 1)
                Month(Pos + 1) = TempMonth
            End If
        Next Pos
    Next Pass
    
    For I = 1 To Ctr
        
        picResults.Print WTO(I), Tab(30); Year(I), Month(I)
    Next I
End Sub


Private Sub cmdSortYear_Click()
    'This button sorts the members by the year that they joined.
    Dim WTO(1 To 1000) As String
    Dim Month(1 To 1000) As String
    Dim Year(1 To 1000) As Single
    Dim TempYear As Single
    Dim TempWTO As String
    Dim TempMonth As String
    Dim Ctr As Integer
    Dim Pos As Integer
    Dim Pass As Integer
    Dim I As Integer
    
    Open App.Path & "\wto.txt" For Input As #1
    
    Ctr = 0
    Do Until EOF(1)
        Ctr = Ctr + 1
        Input #1, WTO(Ctr), Month(Ctr), Year(Ctr)
    Loop
    Close #1
    
    For Pass = 1 To Ctr - 1
        For Pos = 1 To Ctr - Pass
            If Year(Pos) > Year(Pos + 1) Then
                TempYear = Year(Pos)
                Year(Pos) = Year(Pos + 1)
                Year(Pos + 1) = TempYear
                
                TempWTO = WTO(Pos)
                WTO(Pos) = WTO(Pos + 1)
                WTO(Pos + 1) = TempWTO
                
                TempMonth = Month(Pos)
                Month(Pos) = Month(Pos + 1)
                Month(Pos + 1) = TempMonth
            End If
        Next Pos
    Next Pass
    picResults.Cls
    For I = 1 To Ctr
        picResults.Print Year(I), WTO(I), Tab(30); Month(I)
    Next I
                
    
End Sub
