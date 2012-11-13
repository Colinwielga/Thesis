VERSION 5.00
Begin VB.Form frmWrestling 
   BackColor       =   &H00FF0000&
   Caption         =   "SJU Wrestling Stats"
   ClientHeight    =   6540
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10080
   LinkTopic       =   "Form1"
   ScaleHeight     =   6540
   ScaleWidth      =   10080
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pbxAction 
      Height          =   1215
      Left            =   1320
      Picture         =   "FRMWRE~1.frx":0000
      ScaleHeight     =   1155
      ScaleWidth      =   1515
      TabIndex        =   10
      Top             =   0
      Width           =   1575
   End
   Begin VB.TextBox txtHeader 
      BackColor       =   &H000000FF&
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Text            =   "Select Read and then choose category you wish to look up."
      Top             =   1320
      Width           =   4335
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find wrestler's Info"
      Height          =   735
      Left            =   1920
      TabIndex        =   8
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton cmdWP 
      Caption         =   "Compute Winning %"
      Height          =   735
      Left            =   3480
      TabIndex        =   7
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton cmdP 
      Caption         =   "Pins"
      Height          =   735
      Left            =   240
      TabIndex        =   6
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton cmdTD 
      Caption         =   "Takedowns"
      Height          =   735
      Left            =   3480
      TabIndex        =   5
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton cmdW 
      Caption         =   "Wins"
      Height          =   735
      Left            =   1920
      TabIndex        =   4
      Top             =   1800
      Width           =   1455
   End
   Begin VB.PictureBox pbxSJU 
      Height          =   2295
      Left            =   120
      Picture         =   "FRMWRE~1.frx":0D27
      ScaleHeight     =   2235
      ScaleWidth      =   4755
      TabIndex        =   3
      Top             =   4200
      Width           =   4815
   End
   Begin VB.PictureBox pbxResults 
      BackColor       =   &H000000FF&
      FillColor       =   &H00C00000&
      Height          =   6375
      Left            =   5040
      ScaleHeight     =   6315
      ScaleWidth      =   4875
      TabIndex        =   2
      Top             =   120
      Width           =   4935
   End
   Begin VB.CommandButton cmdRead 
      BackColor       =   &H000000FF&
      Caption         =   "Read"
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   1800
      Width           =   1575
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   735
      Left            =   1920
      TabIndex        =   0
      Top             =   3480
      Width           =   1455
   End
End
Attribute VB_Name = "frmWrestling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : ProjWrestlingInfo (Dan Fiecke's VB Project.vbp)
'Form Name : SJU Wrestling Stats (Wrestling.frm)
'Author: Dan Fiecke
'Date Written: October 29, 2003
'Purpose of Form: To get wrestling information on wrestler's performances last wrestling season
                'to compute the winning percentage of all wrestlers
                'to find out individual stats

'Option Explicit is a command to force
'the user to explicitly declare all variables
'before they can be used.
Option Explicit
Dim strName(1 To 16) As String
Dim strPath As String
Dim Takedowns(1 To 16) As Integer
Dim Wins(1 To 16) As Integer
Dim Losses(1 To 16) As Integer
Dim Pins(1 To 16) As Integer
Dim i As Integer

Private Sub cmdFind_Click()
Dim Found As Boolean
Dim iWrestler As String
Dim W As Integer
W = 16
i = 0
'has user enter a name of wrestler to look up in the txt
iWrestler = InputBox("Enter the name of the wrestler you wish to find")
Found = False
'Clear whatever may be in pbxResults for repeated use.
pbxResults.Cls
'Prints header on top of the picture box
pbxResults.Print "Wrestler"; Tab(25); "Wins"; Tab(30); "Losses"; Tab(40); "Takedowns"; Tab(55); "Pins"
Do While i <= W - 1 And Found = False
    'counts the number of times you go through your program
    i = i + 1
    If iWrestler = strName(i) Then
    Found = True
    End If
Loop
If Found = True Then
        'prints out the name of the wrestler you entered and their stats
        pbxResults.Print iWrestler; "'s stats are:"; Tab(25); Wins(i); Tab(31); Losses(i); Tab(44); Takedowns(i); Tab(55); Pins(i)
    Else
        'gives a pop up message and tells them there is no wrestler with that name
        MsgBox ("Did not wrestle for St. John's University")
End If
End Sub

Private Sub cmdP_Click()
Dim N As Integer
Dim pass As Integer
Dim temp As Integer
Dim temps As String
'Clear whatever may be in pbxResults for repeated use.
pbxResults.Cls
N = 16
'prints the header wrestler and pins on the top of the Picture box
pbxResults.Print "Wrestler"; Tab(20); "Pins"
For pass = 1 To N - 1
    For i = 1 To N - pass
        'sorts through the list of wrestlers to see who has the most pins and then
        'put them in order from greatest amount to least amount of pins
        If Pins(i) < Pins(i + 1) Then
            temp = Pins(i + 1)
            Pins(i + 1) = Pins(i)
            Pins(i) = temp
            temps = strName(i + 1)
            strName(i + 1) = strName(i)
            strName(i) = temps
            temp = Wins(i + 1)
            Wins(i + 1) = Wins(i)
            Wins(i) = temp
            temp = Takedowns(i + 1)
            Takedowns(i + 1) = Takedowns(i)
            Takedowns(i) = temp
            
        End If
    Next i
Next pass
For i = 1 To N
    'prints out the results of the if statement
    pbxResults.Print strName(i); Tab(20); Pins(i)
Next i
End Sub

Private Sub cmdQuit_Click()
    End  'Exit out of the program
End Sub

Private Sub cmdRead_Click()
'Loads File into VB project
Open strPath For Input As #1
'Clear whatever may be in pbxResults for repeated use.
pbxResults.Cls
'Prints header on top of pbxResults
pbxResults.Print "Wrestler"; Tab(20); "Wins"; Tab(27); "Losses"; Tab(35); "Takedowns"; Tab(50); "Pins"
'Number of passes through a list
For i = 1 To 16
'Gets all info from txt and loads it back into respected arrays
    Input #1, strName(i), Wins(i), Losses(i), Takedowns(i), Pins(i)
'Prints out all the info from the txt document onto the VB program
    pbxResults.Print strName(i); Tab(20); Wins(i); Tab(27); Losses(i); Tab(37); Takedowns(i); Tab(51); Pins(i)

Next i
End Sub

Private Sub cmdTD_Click()
Dim N As Integer
Dim pass As Integer
Dim temp As Integer
Dim temps As String
'Clear whatever may be in pbxResults for repeated use.
pbxResults.Cls
N = 16
'prints out header on the top of your picture box
pbxResults.Print "Wrestler"; Tab(20); "Takedowns"
For pass = 1 To N - 1
    For i = 1 To N - pass
        If Takedowns(i) < Takedowns(i + 1) Then
            temp = Takedowns(i + 1)
            Takedowns(i + 1) = Takedowns(i)
            Takedowns(i) = temp
            temps = strName(i + 1)
            strName(i + 1) = strName(i)
            strName(i) = temps
            temp = Wins(i + 1)
            Wins(i + 1) = Wins(i)
            Wins(i) = temp
            temp = Pins(i + 1)
            Pins(i + 1) = Pins(i)
            Pins(i) = temp
            
        End If
    Next i
Next pass
For i = 1 To N
    'prints out the results of the if statement
    pbxResults.Print strName(i); Tab(20); Takedowns(i)
Next i
End Sub

Private Sub cmdW_Click()
Dim N As Integer
Dim pass As Integer
Dim temp As Integer
Dim temps As String
'Clear whatever may be in pbxResults for repeated use.
pbxResults.Cls
N = 16
'prints out header on the top of your picture box
pbxResults.Print "Wrestler"; Tab(20); "Wins"
For pass = 1 To N - 1
    For i = 1 To N - pass
        If Wins(i) < Wins(i + 1) Then
            temp = Wins(i + 1)
            Wins(i + 1) = Wins(i)
            Wins(i) = temp
            temps = strName(i + 1)
            strName(i + 1) = strName(i)
            strName(i) = temps
            temp = Pins(i + 1)
            Pins(i + 1) = Pins(i)
            Pins(i) = temp
            temp = Takedowns(i + 1)
            Takedowns(i + 1) = Takedowns(i)
            Takedowns(i) = temp
            
        End If
    Next i
Next pass
For i = 1 To N
    'prints out the results of the if statement
    pbxResults.Print strName(i); Tab(20); Wins(i)
Next i
End Sub

Private Sub cmdWP_Click()
Dim WP(1 To 16) As Single
Dim pass As Integer
Dim N As Single
Dim temp As Single
Dim temps As String
Dim X As Single
'Clear whatever may be in pbxResults for repeated use.
pbxResults.Cls
'prints out header on the top of your picture box
pbxResults.Print "Wrestler"; Tab(20); "Wrestler's Winning % is:"
For N = 1 To 16
        ' formula used for getting the winning percentage
        WP(N) = Wins(N) / (Wins(N) + Losses(N))
Next N
For N = 1 To 16
   'prints out the results of the if statement
   pbxResults.Print strName(N); Tab(30); FormatPercent(WP(N))
Next N
End Sub

Private Sub Form_Load()
strPath = "N:\CS130\djfiecke\Wrestling.txt"
End Sub
