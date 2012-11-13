VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H008080FF&
   Caption         =   "Form3"
   ClientHeight    =   8025
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10500
   LinkTopic       =   "Form3"
   ScaleHeight     =   8025
   ScaleWidth      =   10500
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdnormal 
      Caption         =   "Average Totals for Starters"
      Height          =   855
      Left            =   2520
      TabIndex        =   31
      Top             =   3960
      Width           =   1575
   End
   Begin VB.CommandButton cmdF2 
      Caption         =   "Back to Form 2"
      Height          =   735
      Left            =   2520
      TabIndex        =   30
      Top             =   6000
      Width           =   1575
   End
   Begin VB.CommandButton cmdF1 
      Caption         =   "Back to Form 1"
      Height          =   735
      Left            =   2520
      TabIndex        =   29
      Top             =   5040
      Width           =   1575
   End
   Begin VB.CommandButton cmdQuit3 
      Caption         =   "QUIT"
      Height          =   735
      Left            =   480
      TabIndex        =   28
      Top             =   6000
      Width           =   1575
   End
   Begin VB.CommandButton cmdClear3 
      Caption         =   "CLEAR"
      Height          =   735
      Left            =   480
      TabIndex        =   27
      Top             =   5040
      Width           =   1575
   End
   Begin VB.CommandButton cmdTotals 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Show Totals"
      Height          =   855
      Left            =   480
      TabIndex        =   25
      Top             =   3960
      Width           =   1575
   End
   Begin VB.PictureBox picresults4 
      BackColor       =   &H00FFC0C0&
      Height          =   2895
      Left            =   4680
      ScaleHeight     =   2835
      ScaleWidth      =   3555
      TabIndex        =   24
      Top             =   960
      Width           =   3615
   End
   Begin VB.TextBox txta5 
      Height          =   375
      Left            =   3480
      TabIndex        =   23
      Top             =   3360
      Width           =   615
   End
   Begin VB.TextBox txtr5 
      Height          =   375
      Left            =   2760
      TabIndex        =   22
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox txtp5 
      Height          =   375
      Left            =   2040
      TabIndex        =   21
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox txta4 
      Height          =   375
      Left            =   3480
      TabIndex        =   20
      Top             =   2880
      Width           =   615
   End
   Begin VB.TextBox txtr4 
      Height          =   375
      Left            =   2760
      TabIndex        =   19
      Top             =   2880
      Width           =   495
   End
   Begin VB.TextBox txtp4 
      Height          =   375
      Left            =   2040
      TabIndex        =   18
      Top             =   2880
      Width           =   495
   End
   Begin VB.TextBox txta3 
      Height          =   375
      Left            =   3480
      TabIndex        =   17
      Top             =   2400
      Width           =   615
   End
   Begin VB.TextBox txtr3 
      Height          =   375
      Left            =   2760
      TabIndex        =   16
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox txtp3 
      Height          =   375
      Left            =   2040
      TabIndex        =   15
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox txta2 
      Height          =   375
      Left            =   3480
      TabIndex        =   14
      Top             =   1920
      Width           =   615
   End
   Begin VB.TextBox txtr2 
      Height          =   375
      Left            =   2760
      TabIndex        =   13
      Top             =   1920
      Width           =   495
   End
   Begin VB.TextBox txtp2 
      Height          =   375
      Left            =   2040
      TabIndex        =   12
      Top             =   1920
      Width           =   495
   End
   Begin VB.TextBox txta1 
      Height          =   375
      Left            =   3480
      TabIndex        =   11
      Top             =   1440
      Width           =   615
   End
   Begin VB.TextBox txtr1 
      Height          =   375
      Left            =   2760
      TabIndex        =   10
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox txtp1 
      Height          =   375
      Left            =   2040
      TabIndex        =   9
      Top             =   1440
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   3555
      Left            =   4680
      Picture         =   "Form4.frx":0000
      Top             =   4080
      Width           =   3540
   End
   Begin VB.Label LabelEnter 
      BackColor       =   &H008080FF&
      Caption         =   "Enter Statistics Below      (For one game )"
      Height          =   495
      Left            =   2040
      TabIndex        =   26
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label LabelA 
      BackColor       =   &H0080C0FF&
      Caption         =   "ASTS"
      Height          =   375
      Left            =   3480
      TabIndex        =   8
      Top             =   960
      Width           =   495
   End
   Begin VB.Label LabelR 
      BackColor       =   &H0080C0FF&
      Caption         =   "REBS"
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      Top             =   960
      Width           =   495
   End
   Begin VB.Label LabelPTS 
      BackColor       =   &H0080C0FF&
      Caption         =   "PTS"
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Top             =   960
      Width           =   495
   End
   Begin VB.Label LabPG 
      Caption         =   "D. Jones"
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label LabSG 
      Caption         =   "J.Jackson"
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label LabSF 
      Caption         =   "T. Lewis"
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label LabPF 
      Caption         =   "F. Lowe"
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label LabC 
      Caption         =   "A. Anderson"
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label LabelStarters 
      BackColor       =   &H008080FF&
      Caption         =   "Normal Starters"
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   1200
      Width           =   1335
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name:TeamStats stats.vbg
'Author Matt de Leon
'Form Name: Form 3 (Form4.frm)
'Written March 15 2004
'This form allows user to insert statistics of a starting lineup and compare the stats of one game to the season averages.
Option Explicit
'Declare Variables to be used
Dim Ptotal As Single
Dim RTotal As Single
Dim ATotal As Single
Dim s1 As Single, s2 As Single, s3 As Single, s4 As Single, s5 As Single
Dim r1 As Single, r2 As Single, r3 As Single, r4 As Single, r5 As Single
Dim a1 As Single, a2 As Single, a3 As Single, a4 As Single, a5 As Single





Private Sub cmdClear3_Click()
picresults4.Cls
End Sub

Private Sub cmdF1_Click()
'Going to form 1
Form3.Hide
Form2.Hide
Form1.Show
End Sub

Private Sub cmdF2_Click()
'going to form 2
Form2.Show
Form3.Hide
Form1.Hide

End Sub

Private Sub cmdnormal_Click()
'Displays season Averages of the Starters on team
picresults4.Print "-------------------------Normal Averages--------------------"
picresults4.Print "PTS"; Tab(15); "REBS"; Tab(30); "ASTS"; Tab(40)
picresults4.Print "77.4"; Tab(15); "35.7"; Tab(30); "19.1"; Tab(40)
End Sub

Private Sub cmdQuit3_Click()
End
End Sub

Private Sub cmdTotals_Click()
'Allows user to put in statistics(points,rebounds,assists) from a particular game and calculates the totals.
s1 = txtp1.Text
s2 = txtp2.Text
s3 = txtp3.Text
s4 = txtp4.Text
s5 = txtp5.Text
 Ptotal = s1 + s2 + s3 + s4 + s5
r1 = txtr1.Text
r2 = txtr2.Text
r3 = txtr3.Text
r4 = txtr4.Text
r5 = txtr5.Text
 RTotal = r1 + r2 + r3 + r4 + r5
a1 = txta1.Text
a2 = txta2.Text
a3 = txta3.Text
a4 = txta4.Text
a5 = txta5.Text
 ATotal = a1 + a2 + a3 + a4 + a5
'Displays stat total(Points,Rebounds,Assists)
picresults4.Print "-----------------Starting Lineup Totals-------------------------"
picresults4.Print "PTS"; Tab(15); "REBS"; Tab(30); "ASTS"; Tab(40)
picresults4.Print Ptotal; Tab(15); RTotal; Tab(30); ATotal; Tab(40)
picresults4.Print
'Gives user idea of how the starters compared to their normal averages.
 If Ptotal > 77.4 Then
 picresults4.Print "Above Average Scoring"
 Else
 picresults4.Print "Below Average Scoring"
 End If
 If RTotal > 35.7 Then
 picresults4.Print "Above Average Rebounding"
 Else
 picresults4.Print "Below Average Rebounding"
 End If
If ATotal > 19.1 Then
picresults4.Print "Above Average in Assists"
Else
picresults4.Print "Below Average in Assists"
End If
End Sub










