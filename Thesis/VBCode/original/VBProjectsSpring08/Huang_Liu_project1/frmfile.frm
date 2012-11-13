VERSION 5.00
Begin VB.Form frmfile 
   Caption         =   "Kitty's file"
   ClientHeight    =   6165
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10290
   LinkTopic       =   "Form2"
   Picture         =   "frmfile.frx":0000
   ScaleHeight     =   6165
   ScaleWidth      =   10290
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdsearch 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Find an item Under a prize"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   360
      Picture         =   "frmfile.frx":65CC
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4560
      Width           =   1935
   End
   Begin VB.CommandButton cmddown 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Down"
      Height          =   615
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3720
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdup 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Up"
      Height          =   615
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3720
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdshow 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Show Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   360
      Picture         =   "frmfile.frx":9B7E
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2040
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton cmdinput 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Load Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   360
      Picture         =   "frmfile.frx":144DC
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   360
      Width           =   1935
   End
   Begin VB.PictureBox picoutput 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   3975
      Left            =   2760
      ScaleHeight     =   3915
      ScaleWidth      =   7035
      TabIndex        =   2
      Top             =   360
      Width           =   7095
   End
   Begin VB.CommandButton cmdquit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   8400
      Picture         =   "frmfile.frx":1D396
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton cmdback 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Back to the main page"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   6000
      Picture         =   "frmfile.frx":22388
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4560
      Width           =   2055
   End
End
Attribute VB_Name = "frmfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim NamesArray(1 To 100) As String
Dim Prizes(1 To 100) As Single
Dim Pass As Integer, Pos As Integer
Dim Temp1 As Single
Dim Temp2 As String
Dim CTR As Integer
Dim N As Integer

Private Sub cmdback_Click()
frmfile.Visible = False
frmmain.Visible = True
End Sub

Private Sub cmddown_Click()
picoutput.Cls
    For Pass = 1 To CTR - 1
        For Pos = 1 To CTR - Pass
            If Prizes(Pos) < Prizes(Pos + 1) Then
                Temp1 = Prizes(Pos)
                Prizes(Pos) = Prizes(Pos + 1)
                Prizes(Pos + 1) = Temp1
                Temp2 = NamesArray(Pos)
                NamesArray(Pos) = NamesArray(Pos + 1)
                NamesArray(Pos + 1) = Temp2
            End If
        Next Pos
    Next Pass
picoutput.Print "Data is now in decending sorting."
End Sub

Private Sub cmdinput_Click()
CTR = 0
Open App.Path & "\data.txt" For Input As #1
    Do Until EOF(1)
        ' Get the data from the file
        CTR = CTR + 1
        Input #1, NamesArray(CTR), Prizes(CTR)
    Loop
    Close #1
    
    picoutput.Print "Data Loaded!"
    cmdshow.Visible = True
    cmdup.Visible = True
    cmddown.Visible = True

End Sub

Private Sub cmdquit_Click()
End
End Sub

Private Sub cmdsearch_Click()
    Dim P As Integer
    Dim N1 As Integer
    Dim N2 As Integer
    Dim P1(0 To 100) As Single, Names1(0 To 100) As String
    P = InputBox("What prize do you want to spend on an item?", "prize")
    N2 = 0
    For N1 = 1 To CTR
    
    If P > Prizes(N1) Then
    P1(N2) = Prizes(N1)
    Names1(N2) = NamesArray(N1)
    N2 = N2 + 1
    
    End If
    
    Next N1
    
    MsgBox N2 & "items found", , "Items"
    picoutput.Cls
    picoutput.Print "Name of Item"; Tab(50); "Prize of Item"
    For N = 1 To N2
    picoutput.Print Names1(N); Tab(50); FormatCurrency(P1(N))
    Next N
End Sub

Private Sub cmdshow_Click()

    picoutput.Cls
    
    picoutput.Print "Name of Item"; Tab(50); "Prize of Item"
    For N = 1 To CTR
        picoutput.Print NamesArray(N); Tab(50); FormatCurrency(Prizes(N))
    Next N

End Sub

Private Sub cmdup_Click()
picoutput.Cls
    For Pass = 1 To CTR - 1
        For Pos = 1 To CTR - Pass
            If Prizes(Pos) > Prizes(Pos + 1) Then
                Temp1 = Prizes(Pos)
                Prizes(Pos) = Prizes(Pos + 1)
                Prizes(Pos + 1) = Temp1
                Temp2 = NamesArray(Pos)
                NamesArray(Pos) = NamesArray(Pos + 1)
                NamesArray(Pos + 1) = Temp2
            End If
        Next Pos
    Next Pass
    picoutput.Print "Data is now in decending sorting."
End Sub
