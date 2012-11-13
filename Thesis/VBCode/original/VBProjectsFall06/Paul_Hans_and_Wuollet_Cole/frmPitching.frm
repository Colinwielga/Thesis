VERSION 5.00
Begin VB.Form frmPitching 
   BackColor       =   &H00FF0000&
   Caption         =   "Pitching Stats"
   ClientHeight    =   7380
   ClientLeft      =   2055
   ClientTop       =   3090
   ClientWidth     =   9855
   FillColor       =   &H00FF0000&
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7380
   ScaleWidth      =   9855
   Visible         =   0   'False
   Begin VB.CommandButton cmdreturn 
      Caption         =   "Return To Main Menu"
      Height          =   615
      Left            =   3960
      TabIndex        =   10
      Top             =   6600
      Width           =   3015
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   6255
      Left            =   2520
      ScaleHeight     =   6195
      ScaleWidth      =   6795
      TabIndex        =   9
      Top             =   120
      Width           =   6855
   End
   Begin VB.CommandButton cmdIP 
      Caption         =   "Sort By Innings Pitched"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   5880
      Width           =   1935
   End
   Begin VB.CommandButton cmdCG 
      Caption         =   "Sort By Complete Games"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   5160
      Width           =   1935
   End
   Begin VB.CommandButton cmdSV 
      Caption         =   "Sort By Saves"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   4440
      Width           =   1935
   End
   Begin VB.CommandButton cmdLosses 
      Caption         =   "Sort By Losses"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   3720
      Width           =   1935
   End
   Begin VB.CommandButton cmdW 
      Caption         =   "Sort By Wins"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   3000
      Width           =   1935
   End
   Begin VB.CommandButton cmdERA 
      Caption         =   "Sort By ERA"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   1935
   End
   Begin VB.CommandButton cmdGame 
      Caption         =   "Sort By Games Pitched"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   1935
   End
   Begin VB.CommandButton cmdnam 
      Caption         =   "Sort By Name"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1935
   End
   Begin VB.CommandButton cmdclick 
      Caption         =   "Load Data (Click Here First)"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmPitching"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: 1987 World Series
'Form Name: frmPitching
'Authors: Hans Paul and Cole Wuollet
'Date Written: Tuesday November 31, 2006
'Objective: Load and Sort an array from a file, and display different results
Option Explicit
Private Sub cmdCG_Click()
    For Pass = 1 To (Size - 1)
        For Counter = 1 To (Size - Pass)
            If CGs(Counter) < CGs(Counter + 1) Then
                Tempnam = Nams(Counter)
                Nams(Counter) = Nams(Counter + 1)
                Nams(Counter + 1) = Tempnam
            
                Tempcg = CGs(Counter)
                CGs(Counter) = CGs(Counter + 1)
                CGs(Counter + 1) = Tempcg
            End If
        Next Counter
    Next Pass
    picResults.Cls
    picResults.Print "Name"; , , , "Complete Games"
    picResults.Print "***********************************************************"
    For Pos = 1 To Size
        picResults.Print Nams(Pos); Tab; CGs(Pos)
    Next Pos
End Sub

Private Sub cmdERA_Click()
    For Pass = 1 To (Size - 1)
        For Counter = 1 To (Size - Pass)
            If ERAs(Counter) > ERAs(Counter + 1) Then
                Tempnam = Nams(Counter)
                Nams(Counter) = Nams(Counter + 1)
                Nams(Counter + 1) = Tempnam

                Tempe = ERAs(Counter)
                ERAs(Counter) = ERAs(Counter + 1)
                ERAs(Counter + 1) = Tempe
            End If
        Next Counter
    Next Pass
    picResults.Cls
    picResults.Print "Name"; , , , "Earned Run Average"
    picResults.Print "***********************************************************"
    For Pos = 1 To Size
        picResults.Print Nams(Pos); Tab; ERAs(Pos)
    Next Pos
End Sub

Private Sub cmdGame_Click()
    For Pass = 1 To (Size - 1)
        For Counter = 1 To (Size - Pass)
            If Gs(Counter) < Gs(Counter + 1) Then
                Tempnam = Nams(Counter)
                Nams(Counter) = Nams(Counter + 1)
                Nams(Counter + 1) = Tempnam

                Tempgs = Gs(Counter)
                Gs(Counter) = Gs(Counter + 1)
                Gs(Counter + 1) = Tempgs
            End If
        Next Counter
    Next Pass
    picResults.Cls
    picResults.Print "Name"; , , , "Games Pitched"
    picResults.Print "***********************************************************"
    For Pos = 1 To Size
        picResults.Print Nams(Pos); Tab; Gs(Pos)
    Next Pos
End Sub

Private Sub cmdIP_Click()
    For Pass = 1 To (Size - 1)
        For Counter = 1 To (Size - Pass)
            If IPs(Counter) < IPs(Counter + 1) Then
                Tempnam = Nams(Counter)
                Nams(Counter) = Nams(Counter + 1)
                Nams(Counter + 1) = Tempnam

                Tempip = IPs(Counter)
                IPs(Counter) = IPs(Counter + 1)
                IPs(Counter + 1) = Tempip
            End If
        Next Counter
    Next Pass
    picResults.Cls
    picResults.Print "Name"; , , , "Innings Pitched"
    picResults.Print "***********************************************************"
    For Pos = 1 To Size
        picResults.Print Nams(Pos); Tab; IPs(Pos)
    Next Pos
End Sub

Private Sub cmdLosses_Click()
    For Pass = 1 To (Size - 1)
        For Counter = 1 To (Size - Pass)
            If Ls(Counter) < Ls(Counter + 1) Then
                Tempnam = Nams(Counter)
                Nams(Counter) = Nams(Counter + 1)
                Nams(Counter + 1) = Tempnam

                Templ = Ls(Counter)
                Ls(Counter) = Ls(Counter + 1)
                Ls(Counter + 1) = Templ
            End If
        Next Counter
    Next Pass
    picResults.Cls
    picResults.Print "Name"; , , , "Losses"
    picResults.Print "***********************************************************"
    For Pos = 1 To Size
        picResults.Print Nams(Pos); Tab; Ls(Pos)
    Next Pos
End Sub

Private Sub cmdnam_Click()
    For Pass = 1 To (Size - 1)
        For Counter = 1 To (Size - Pass)
            If Nams(Counter) > Nams(Counter + 1) Then
                Tempnam = Nams(Counter)
                Nams(Counter) = Nams(Counter + 1)
                Nams(Counter + 1) = Tempnam

                Tempgs = Gs(Counter)
                Gs(Counter) = Gs(Counter + 1)
                Gs(Counter + 1) = Tempgs
            End If
        Next Counter
    Next Pass
    picResults.Cls
    picResults.Print "Name"; , , , "Games Pitched"
    picResults.Print "***********************************************************"
    For Pos = 1 To Size
        picResults.Print Nams(Pos); Tab; Gs(Pos)
    Next Pos
End Sub

Private Sub cmdReturn_Click()
    frmPitching.Hide
    frmTwins.Show
End Sub

Private Sub cmdSV_Click()
    For Pass = 1 To (Size - 1)
        For Counter = 1 To (Size - Pass)
            If SVs(Counter) < SVs(Counter + 1) Then
                Tempnam = Nams(Counter)
                Nams(Counter) = Nams(Counter + 1)
                Nams(Counter + 1) = Tempnam

                Tempsv = SVs(Counter)
                SVs(Counter) = SVs(Counter + 1)
                SVs(Counter + 1) = Tempsv
            End If
        Next Counter
    Next Pass
    picResults.Cls
    picResults.Print "Name"; , , , "Saves"
    picResults.Print "***********************************************************"
    For Pos = 1 To Size
        picResults.Print Nams(Pos); Tab; SVs(Pos)
    Next Pos
End Sub

Private Sub cmdW_Click()
    For Pass = 1 To (Size - 1)
        For Counter = 1 To (Size - Pass)
            If Ws(Counter) < Ws(Counter + 1) Then
                Tempnam = Nams(Counter)
                Nams(Counter) = Nams(Counter + 1)
                Nams(Counter + 1) = Tempnam

                Tempw = Ws(Counter)
                Ws(Counter) = Ws(Counter + 1)
                Ws(Counter + 1) = Tempw
            End If
        Next Counter
    Next Pass
    picResults.Cls
    picResults.Print "Name"; , , , "Wins"
    picResults.Print "***********************************************************"
    For Pos = 1 To Size
        picResults.Print Nams(Pos); Tab; Ws(Pos)
    Next Pos
End Sub

Private Sub cmdclick_Click()
    Open App.Path & "\Pitchers.txt" For Input As #2
    Counter = 0
    Do Until EOF(2)
        Counter = Counter + 1
        Input #2, Nams(Counter), Gs(Counter), ERAs(Counter), Ws(Counter), Ls(Counter), SVs(Counter), CGs(Counter), IPs(Counter)
    Loop
    Size = Counter
    Close #2
End Sub
