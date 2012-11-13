VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00800000&
   Caption         =   "Form1"
   ClientHeight    =   9135
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   Picture         =   "Twins.frx":0000
   ScaleHeight     =   9135
   ScaleWidth      =   9600
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   1935
      Left            =   0
      Picture         =   "Twins.frx":98E6
      ScaleHeight     =   1875
      ScaleWidth      =   2955
      TabIndex        =   10
      Top             =   7200
      Width           =   3015
      Begin VB.PictureBox Picture2 
         Height          =   1815
         Left            =   3000
         ScaleHeight     =   1815
         ScaleWidth      =   15
         TabIndex        =   11
         Top             =   0
         Width           =   15
      End
   End
   Begin VB.PictureBox pbxresults 
      BackColor       =   &H000000FF&
      Height          =   9975
      Left            =   4680
      ScaleHeight     =   9915
      ScaleWidth      =   5115
      TabIndex        =   0
      Top             =   0
      Width           =   5175
   End
   Begin VB.CommandButton cmdTotal 
      Caption         =   "Designated Hitter (View Total Number of Fans to Watch the Twins since 1961)"
      Height          =   1335
      Left            =   2520
      TabIndex        =   9
      Top             =   5160
      Width           =   2055
   End
   Begin VB.PictureBox oldtwins 
      Height          =   2295
      Left            =   0
      Picture         =   "Twins.frx":B766
      ScaleHeight     =   2235
      ScaleWidth      =   1875
      TabIndex        =   8
      Top             =   4920
      Width           =   1935
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000000FF&
      Caption         =   "Home Plate (Quit)"
      Height          =   855
      Left            =   1680
      MaskColor       =   &H000000FF&
      TabIndex        =   7
      Top             =   3840
      Width           =   1335
   End
   Begin VB.CommandButton cmdSing 
      Caption         =   "Left Field (Find Record For Single Year)"
      Height          =   855
      Left            =   120
      TabIndex        =   6
      Top             =   360
      Width           =   1335
   End
   Begin VB.CommandButton cmdDiv 
      Caption         =   "Centerfield (Show Record in Years they won the Division)"
      Height          =   975
      Left            =   1680
      TabIndex        =   5
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton cmdAtt 
      Caption         =   "Right Field (Sort by Attendance In A Season)"
      Height          =   855
      Left            =   3240
      TabIndex        =   4
      Top             =   360
      Width           =   1335
   End
   Begin VB.CommandButton cmdLosses 
      Caption         =   "3rd Base (Sort by Most Losses In A Season)"
      Height          =   855
      Left            =   480
      TabIndex        =   3
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton cmdWins 
      Caption         =   "2nd Base (Sort By Most Wins In a Season)"
      Height          =   855
      Left            =   1680
      TabIndex        =   2
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton cmdRead 
      BackColor       =   &H000000FF&
      Caption         =   "1st Base (Load Twins Records Year-by-Year) "
      Height          =   855
      Left            =   2760
      MaskColor       =   &H00E0E0E0&
      TabIndex        =   1
      Top             =   2760
      Width           =   1335
   End
   Begin VB.PictureBox Picture3 
      Height          =   1935
      Left            =   2280
      Picture         =   "Twins.frx":10709
      ScaleHeight     =   1875
      ScaleWidth      =   2475
      TabIndex        =   12
      Top             =   7200
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Minnesota Twins (MNTwins.vbp)
'Form Name: Twins Year-By-Year (Twins.frm)
'Brett Hemmelgarn
'Oct. 20th, 2003
'Purpose: This program is to let the user look at the Minnesota Twins
    'year by year and sort the records by wins, losses, attendance,
    'Division Championships and record in a single year
Option Explicit
Dim i As Integer
Dim Year(1 To 42) As Integer
Dim W(1 To 42) As Integer
Dim L(1 To 42) As Integer
Dim Att(1 To 42) As Long
Public Path As String






Private Sub cmdAtt_Click()
Dim Pass As Integer
Dim Temp1 As Long
Dim Temp2 As Integer
Dim Temp3 As Integer
Dim Temp4 As Integer
    'Sorts by attendance and prints from most to least to see
    'which years they had the most attendance and relate it to their W-L
pbxresults.Cls
pbxresults.Print "Year"; Tab(15); "Wins", "Losses", "Attendance"
pbxresults.Print "**************************************************************"
    For Pass = 1 To 41
        For i = 1 To 42 - Pass
            If Att(i) < Att(i + 1) Then
                Temp1 = Att(i)
                Temp2 = W(i)
                Temp3 = L(i)
                Temp4 = Year(i)
                Att(i) = Att(i + 1)
                W(i) = W(i + 1)
                L(i) = L(i + 1)
                Year(i) = Year(i + 1)
                Att(i + 1) = Temp1
                W(i + 1) = Temp2
                L(i + 1) = Temp3
                Year(i + 1) = Temp4
            End If
        Next i
    Next Pass

For i = 1 To 41
    pbxresults.Print Year(i), W(i), L(i), FormatNumber(Att(i), 0)
Next i
    
End Sub

Private Sub cmdDiv_Click()
pbxresults.Cls
pbxresults.Print "Year"; Tab(15); "Wins", "Losses", "Attendance"
pbxresults.Print "**************************************************************"
    'Prints out the year, W-L, and att. of years they won the pennant
Open Path & "Champs.txt" For Input As #2
   For i = 1 To 6
        Input #2, Year(i), W(i), L(i), Att(i)
    Next i
Close #2

For i = 1 To 6
 pbxresults.Print Year(i), W(i), L(i), Att(i)
Next i

End Sub

Private Sub cmdLosses_Click()
Dim Pass As Integer
Dim Temp1 As Integer
Dim Temp2 As Integer
Dim Temp3 As Integer
    'Sorts records and prints from Most losses to least
pbxresults.Cls
pbxresults.Print "Year"; Tab(15); "Wins", "Losses"
pbxresults.Print "***********************************************"
    
    For Pass = 1 To 41
        For i = 1 To 42 - Pass
            If L(i) < L(i + 1) Then
            Temp1 = L(i)
            Temp2 = W(i)
            Temp3 = Year(i)
            L(i) = L(i + 1)
            W(i) = W(i + 1)
            Year(i) = Year(i + 1)
            L(i + 1) = Temp1
            W(i + 1) = Temp2
            Year(i + 1) = Temp3
        End If
    Next i
Next Pass

For i = 1 To 41
    pbxresults.Print Year(i), W(i), L(i)
Next i
    
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdRead_Click()
pbxresults.Cls
pbxresults.Print "Year"; Tab(15); "Wins", "Losses"
pbxresults.Print "***********************************************"
    'Opens the file to use for the rest of the program
Open Path & "Twins.txt" For Input As #1
    For i = 1 To 42
        Input #1, Year(i), W(i), L(i), Att(i)
    Next i
Close #1

For i = 1 To 42
    pbxresults.Print Year(i), W(i), L(i)
Next i

End Sub


Private Sub cmdSing_Click()
Dim One As Integer
Dim Found As Boolean
    'Provides the user with the capability to look for a record in
    'a single year and then print it out individually
pbxresults.Cls
pbxresults.Print "Year"; Tab(15); "Wins", "Losses"
pbxresults.Print "***********************************************"

Found = False
i = 0
    One = InputBox("Enter year from 1961 to 2002", "Twins W-L Record")
    
    Do Until i = 42 And Found = False
        i = i + 1
            If One = Year(i) Then
             Found = True
            Exit Do
        End If
    Loop

If Found = True Then
    pbxresults.Print Year(i), W(i), L(i)
Else
    MsgBox "Sorry, you must enter a year from 1961-2002", , "Error"
End If

End Sub

Private Sub cmdTotal_Click()
Dim Total As Long
    'Adds up attendance to give total figure of att. throughout history
pbxresults.Cls
    For i = 1 To 42
        Total = Total + Att(i)
    Next i
    
    pbxresults.Print FormatNumber((Total), 0), "people have watched the Twins play over the years."
End Sub

Private Sub cmdWins_Click()
Dim Pass As Integer
Dim Temp1 As Integer
Dim Temp2 As Integer
Dim Temp3 As Integer
    'Sorts record by wins and then prints from most wins to least
pbxresults.Cls
pbxresults.Print "Year"; Tab(15); "Wins", "Losses"
pbxresults.Print "***********************************************"
    For Pass = 1 To 41
        For i = 1 To 42 - Pass
            If W(i) < W(i + 1) Then
            Temp1 = W(i)
            Temp2 = L(i)
            Temp3 = Year(i)
            W(i) = W(i + 1)
            L(i) = L(i + 1)
            Year(i) = Year(i + 1)
            W(i + 1) = Temp1
            L(i + 1) = Temp2
            Year(i + 1) = Temp3
        End If
    Next i
Next Pass

For i = 1 To 41
    pbxresults.Print Year(i), W(i), L(i)
Next i
End Sub

Private Sub VScroll1_Change()

End Sub

Private Sub Form_Load()
Path = "N:\CS130\handin\project\"
End Sub
