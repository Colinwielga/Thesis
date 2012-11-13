VERSION 5.00
Begin VB.Form frmRoster 
   BackColor       =   &H0000FFFF&
   Caption         =   "Team Roster"
   ClientHeight    =   7455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10080
   LinkTopic       =   "Form1"
   ScaleHeight     =   7455
   ScaleWidth      =   10080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSortHeight 
      BackColor       =   &H0000C000&
      Caption         =   "Clear names and sort again"
      Height          =   1095
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdListOption 
      BackColor       =   &H0000C000&
      Caption         =   "See Roster"
      Height          =   1335
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   240
      Width           =   1575
   End
   Begin VB.OptionButton optHeight 
      BackColor       =   &H0000C000&
      Caption         =   "Sort players by height"
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   1200
      Width           =   2655
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H0000FFFF&
      Caption         =   "Back to main page."
      Height          =   1215
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5520
      Width           =   2415
   End
   Begin VB.OptionButton optNumerical 
      BackColor       =   &H0000C000&
      Caption         =   "Sort players numerically"
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   240
      Width           =   2655
   End
   Begin VB.PictureBox picRoster 
      Height          =   3855
      Left            =   240
      Picture         =   "frmRoster.frx":0000
      ScaleHeight     =   3795
      ScaleWidth      =   4275
      TabIndex        =   1
      Top             =   1920
      Width           =   4335
   End
   Begin VB.Label lblCoach 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Head Coach Joe Swecker"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   0
      Top             =   4560
      Width           =   3855
   End
   Begin VB.Image Image1 
      Height          =   6795
      Left            =   5040
      Picture         =   "frmRoster.frx":3FA8
      Top             =   240
      Width           =   5040
   End
End
Attribute VB_Name = "frmRoster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBack_Click()
frmStormBball.Show
frmRoster.Hide
End Sub

Private Sub cmdListOption_Click() 'This button can sort roster by number or height
Dim Names(1 To 100) As String, Numbers(1 To 100) As Integer, Height(1 To 100) As Single
Dim CTR As Integer, Pos As Integer
Dim Temp1 As Single, Temp2 As String, Temp3 As Integer
Dim Pass As Integer, I As Integer

Open App.Path & "\roster.txt" For Input As #1 'Opens array of names, numbers, heights

CTR = 0

picRoster.Print "Name"; Tab(20), "Number"; Tab(40), "Height"

If optNumerical.Value = True Then 'controls which option button is selected
    Do Until EOF(1) 'Enters the entire array
        CTR = CTR + 1
        Input #1, Names(CTR), Numbers(CTR), Height(CTR)
    Loop
    Close #1

    For Pass = 1 To CTR - 1 'This series sorts the array
        For Pos = 1 To CTR - Pass
            If Numbers(Pos) > Numbers(Pos + 1) Then
                Temp1 = Numbers(Pos) 'starts with number as first temp
                Numbers(Pos) = Numbers(Pos + 1) 'becuase this is what i want to sort by
                Numbers(Pos + 1) = Temp1
            
                Temp2 = Names(Pos)
                Names(Pos) = Names(Pos + 1)
                Names(Pos + 1) = Temp2

                Temp3 = Height(Pos)
                Height(Pos) = Height(Pos + 1)
                Height(Pos + 1) = Temp3
            End If
        Next Pos
    Next Pass
For I = 1 To CTR
picRoster.Print Names(I); Tab(20), Numbers(I); Tab(40), Height(I)
Next I 'prints the roster sorted by number
End If


If optHeight.Value = True Then 'controls the option button
    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, Names(CTR), Numbers(CTR), Height(CTR)
    Loop
    Close #1
    
    For Pass = 1 To CTR - 1
        For Pos = 1 To CTR - Pass
            If Height(Pos) > Height(Pos + 1) Then
                Temp1 = Height(Pos) 'first temp is height
                Height(Pos) = Height(Pos + 1) 'sorting by height
                Height(Pos + 1) = Temp1
            
                Temp2 = Names(Pos)
                Names(Pos) = Names(Pos + 1)
                Names(Pos + 1) = Temp2

                Temp3 = Numbers(Pos)
                Numbers(Pos) = Numbers(Pos + 1)
                Numbers(Pos + 1) = Temp3
            End If
        Next Pos
    Next Pass
For I = 1 To CTR
picRoster.Print Names(I); Tab(20), Numbers(I); Tab(40), (Height(I))
Next I
End If
cmdSortHeight.Visible = True

End Sub 'http://www.samspublishing.com/library/content.asp?b=STY_VB6_24hours&seqNum=116&rl=1

Private Sub Image2_Click()

End Sub

Private Sub cmdSortHeight_Click()
picRoster.Cls
optNumerical.Value = False
optHeight.Value = False
End Sub


