VERSION 5.00
Begin VB.Form frmSortTwins 
   BackColor       =   &H000000FF&
   Caption         =   "Rank the Twins Players"
   ClientHeight    =   8220
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14790
   LinkTopic       =   "Form1"
   Picture         =   "frmSortTwins.frx":0000
   ScaleHeight     =   8220
   ScaleWidth      =   14790
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   615
      Left            =   960
      TabIndex        =   6
      Top             =   6240
      Width           =   1695
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back to Main Menu"
      Height          =   615
      Left            =   960
      TabIndex        =   5
      Top             =   5400
      Width           =   1695
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFFFF&
      Height          =   5655
      Left            =   3000
      ScaleHeight     =   5595
      ScaleWidth      =   7275
      TabIndex        =   4
      Top             =   600
      Width           =   7335
   End
   Begin VB.CommandButton cmdSortOBP 
      Caption         =   "Sort by OBP"
      Height          =   735
      Left            =   960
      TabIndex        =   3
      Top             =   3480
      Width           =   1695
   End
   Begin VB.CommandButton cmdSortSLG 
      Caption         =   "Sort by SLG"
      Height          =   735
      Left            =   960
      TabIndex        =   2
      Top             =   2520
      Width           =   1695
   End
   Begin VB.CommandButton cmdSortOPS 
      Caption         =   "Sort by OPS"
      Height          =   735
      Left            =   960
      TabIndex        =   1
      Top             =   4440
      Width           =   1695
   End
   Begin VB.CommandButton cmdSortBA 
      Caption         =   "Sort by BA"
      Height          =   735
      Left            =   960
      TabIndex        =   0
      Top             =   1560
      Width           =   1695
   End
End
Attribute VB_Name = "frmSortTwins"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Pass As Integer
Dim Pos As Integer
Dim Temp1 As String
Dim Temp2 As Single
Dim J As Integer

'Baseball Batting Statistics
'frmSortTwins
'Aaron Walsh
'March 24, 2009
'This program will sort from greatest to least various batting statistics like BA, OPS, OBP, and SLG
'for Twins players based on 2008 numbers in certain batting catagories

Private Sub cmdBack_Click()
    frmSortTwins.Hide
    frmInitialform.Show
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdSortBA_Click()
'this sorts the batting avg from highest to lowest and prints it
    For Pass = 1 To 8
        For Pos = 1 To 9 - Pass
            If BA(Pos) < BA(Pos + 1) Then
                Temp1 = TwinsNames(Pos)
                TwinsNames(Pos) = TwinsNames(Pos + 1)
                TwinsNames(Pos + 1) = Temp1
                Temp2 = BA(Pos)
                BA(Pos) = BA(Pos + 1)
                BA(Pos + 1) = Temp2
                Temp2 = OBP(Pos)
                OBP(Pos) = OBP(Pos + 1)
                OBP(Pos + 1) = Temp2
                Temp2 = OPS(Pos)
                OPS(Pos) = OPS(Pos + 1)
                OPS(Pos + 1) = Temp2
                Temp2 = SLG(Pos)
                SLG(Pos) = SLG(Pos + 1)
                SLG(Pos + 1) = Temp2
            End If
        Next Pos
    Next Pass
    picResults.Cls
    picResults.Print "Player", "Batting Average"
    picResults.Print "*************************************"
    For J = 1 To 9
        picResults.Print TwinsNames(J), FormatNumber(BA(J), 3)
    Next J
End Sub

Private Sub cmdSortOBP_Click()
'this sorts the on base percentage from highest to lowest and prints it
    For Pass = 1 To 8
        For Pos = 1 To 9 - Pass
            If OBP(Pos) < OBP(Pos + 1) Then
                Temp1 = TwinsNames(Pos)
                TwinsNames(Pos) = TwinsNames(Pos + 1)
                TwinsNames(Pos + 1) = Temp1
                Temp2 = BA(Pos)
                BA(Pos) = BA(Pos + 1)
                BA(Pos + 1) = Temp2
                Temp2 = OBP(Pos)
                OBP(Pos) = OBP(Pos + 1)
                OBP(Pos + 1) = Temp2
                Temp2 = OPS(Pos)
                OPS(Pos) = OPS(Pos + 1)
                OPS(Pos + 1) = Temp2
                Temp2 = SLG(Pos)
                SLG(Pos) = SLG(Pos + 1)
                SLG(Pos + 1) = Temp2
            End If
        Next Pos
    Next Pass
    picResults.Cls
    picResults.Print "Player", "On Base Percentage"
    picResults.Print "****************************************"
    For J = 1 To 9
        picResults.Print TwinsNames(J), FormatNumber(OBP(J), 3)
    Next J
End Sub

Private Sub cmdSortOPS_Click()
'this sorts the OPS from highest to lowest and prints it
    For Pass = 1 To 8
        For Pos = 1 To 9 - Pass
            If OPS(Pos) < OPS(Pos + 1) Then
                Temp1 = TwinsNames(Pos)
                TwinsNames(Pos) = TwinsNames(Pos + 1)
                TwinsNames(Pos + 1) = Temp1
                Temp2 = BA(Pos)
                BA(Pos) = BA(Pos + 1)
                BA(Pos + 1) = Temp2
                Temp2 = OBP(Pos)
                OBP(Pos) = OBP(Pos + 1)
                OBP(Pos + 1) = Temp2
                Temp2 = OPS(Pos)
                OPS(Pos) = OPS(Pos + 1)
                OPS(Pos + 1) = Temp2
                Temp2 = SLG(Pos)
                SLG(Pos) = SLG(Pos + 1)
                SLG(Pos + 1) = Temp2
            End If
        Next Pos
    Next Pass
    picResults.Cls
    picResults.Print "Player", "On Base plus Slugging"
    picResults.Print "********************************************"
    For J = 1 To 9
        picResults.Print TwinsNames(J), FormatNumber(OPS(J), 3)
    Next J
End Sub

Private Sub cmdSortSLG_Click()
'this sorts the slugging % from highest to lowest and prints it
    For Pass = 1 To 8
        For Pos = 1 To 9 - Pass
            If SLG(Pos) < SLG(Pos + 1) Then
                Temp1 = TwinsNames(Pos)
                TwinsNames(Pos) = TwinsNames(Pos + 1)
                TwinsNames(Pos + 1) = Temp1
                Temp2 = BA(Pos)
                BA(Pos) = BA(Pos + 1)
                BA(Pos + 1) = Temp2
                Temp2 = OBP(Pos)
                OBP(Pos) = OBP(Pos + 1)
                OBP(Pos + 1) = Temp2
                Temp2 = OPS(Pos)
                OPS(Pos) = OPS(Pos + 1)
                OPS(Pos + 1) = Temp2
                Temp2 = SLG(Pos)
                SLG(Pos) = SLG(Pos + 1)
                SLG(Pos + 1) = Temp2
            End If
        Next Pos
    Next Pass
    picResults.Cls
    picResults.Print "Player", "Slugging Percentage"
    picResults.Print "********************************************"
    For J = 1 To 9
        picResults.Print TwinsNames(J), FormatNumber(SLG(J), 3)
    Next J
End Sub

