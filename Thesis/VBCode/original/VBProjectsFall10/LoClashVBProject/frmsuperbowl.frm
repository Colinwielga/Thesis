VERSION 5.00
Begin VB.Form frmsuperbowl 
   BackColor       =   &H000080FF&
   Caption         =   "superbowl"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   15180
   ScaleWidth      =   24960
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picwin 
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6735
      Left            =   840
      ScaleHeight     =   6675
      ScaleWidth      =   5115
      TabIndex        =   22
      Top             =   4560
      Width           =   5175
   End
   Begin VB.CommandButton cmdwin 
      Caption         =   "Winners"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   21
      Top             =   3960
      Width           =   3015
   End
   Begin VB.PictureBox picresults2 
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   14880
      ScaleHeight     =   795
      ScaleWidth      =   4395
      TabIndex        =   18
      Top             =   1920
      Width           =   4455
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   13440
      TabIndex        =   17
      Top             =   6720
      Width           =   1695
   End
   Begin VB.PictureBox picbowl 
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   17880
      ScaleHeight     =   675
      ScaleWidth      =   1395
      TabIndex        =   16
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton cmdposition 
      Caption         =   "MVPs By Position"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4920
      TabIndex        =   11
      Top             =   11760
      Width           =   2055
   End
   Begin VB.CommandButton cmdmvp 
      Caption         =   "Super Bowl MVPs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2760
      TabIndex        =   10
      Top             =   11760
      Width           =   1935
   End
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   840
      TabIndex        =   9
      Top             =   11760
      Width           =   1695
   End
   Begin VB.PictureBox picchampions 
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   9240
      ScaleHeight     =   1155
      ScaleWidth      =   10035
      TabIndex        =   8
      Top             =   10080
      Width           =   10095
   End
   Begin VB.TextBox txtyear 
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   11400
      TabIndex        =   7
      Top             =   9120
      Width           =   1455
   End
   Begin VB.PictureBox picteams2 
      BackColor       =   &H0080C0FF&
      Height          =   1815
      Left            =   16920
      ScaleHeight     =   1755
      ScaleWidth      =   2355
      TabIndex        =   6
      Top             =   4440
      Width           =   2415
   End
   Begin VB.PictureBox picteams 
      BackColor       =   &H0080C0FF&
      Height          =   1815
      Left            =   9240
      ScaleHeight     =   1755
      ScaleWidth      =   2355
      TabIndex        =   5
      Top             =   4440
      Width           =   2415
   End
   Begin VB.PictureBox picnewyear 
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6720
      ScaleHeight     =   675
      ScaleWidth      =   1515
      TabIndex        =   4
      Top             =   2880
      Width           =   1575
   End
   Begin VB.PictureBox piclocation 
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10440
      ScaleHeight     =   675
      ScaleWidth      =   6075
      TabIndex        =   3
      Top             =   3000
      Width           =   6135
   End
   Begin VB.PictureBox picresults 
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9240
      ScaleHeight     =   795
      ScaleWidth      =   4395
      TabIndex        =   2
      Top             =   1920
      Width           =   4455
   End
   Begin VB.CommandButton cmdchampions 
      Caption         =   "Super Bowl Champions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   13080
      TabIndex        =   1
      Top             =   9120
      Width           =   2895
   End
   Begin VB.CommandButton cmdsuperbowl 
      Caption         =   "Super Bowl"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6000
      TabIndex        =   0
      Top             =   1920
      Width           =   3015
   End
   Begin VB.Label lbltitle 
      BackColor       =   &H000080FF&
      Caption         =   "Super Bowls 1990-2010"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   10080
      TabIndex        =   20
      Top             =   240
      Width           =   8415
   End
   Begin VB.Label lblsv 
      BackColor       =   &H000080FF&
      Caption         =   "V.S."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13920
      TabIndex        =   19
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label lblenter 
      BackColor       =   &H000080FF&
      Caption         =   "Enter Super Bowl"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9240
      TabIndex        =   15
      Top             =   9360
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H000080FF&
      Caption         =   "V.S."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   13800
      TabIndex        =   14
      Top             =   5160
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H000080FF&
      Caption         =   "Super Bowl"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   16920
      TabIndex        =   13
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label lbllocation 
      BackColor       =   &H000080FF&
      Caption         =   "Location"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9360
      TabIndex        =   12
      Top             =   3000
      Width           =   1095
   End
End
Attribute VB_Name = "frmsuperbowl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdchampions_Click()
'identifying variables
Dim teamw(1 To 21) As String
Dim teaml(1 To 21) As String
Dim scored(1 To 21) As String
Dim digits(1 To 21) As Single
Dim place(1 To 21) As String
Dim bowl(1 To 21) As String
Dim P As Integer
Dim bowlyear As String
Dim picyear As Single

'setting what found equals to
found = False

'clear picture box
picchampions.Cls

bowlyear = txtyear.Text


'opens/access the data file
Open App.Path & "\champs.txt" For Input As #1

'filling an array from the specific data file
Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, digits(ctr), teamw(ctr), teaml(ctr), scored(ctr), place(ctr), bowl(ctr)
Loop

'searches for the data until found or at the end of the list
Do While ((Not found) And (P < ctr))
    P = P + 1
    If bowlyear = bowl(P) Then
        found = True
    End If
Loop
'print results based on variable value
If (Not found) Then
        picchampions.Print "Information is Not Available"
    Else
        picchampions.Print "The "; teamw(P); " defeated the "; teaml(P); " by the final of "; scored(P);
End If

'close data file to be re-read
Close
End Sub


Private Sub cmdclear_Click()
'clear picture box
picresults2.Cls
picresults.Cls
piclocation.Cls
picnewyear.Cls
picchampions.Cls
picbowl.Cls
End Sub

Private Sub cmdmvp_Click()
'switch between forms
frmmvp.Show
frmsuperbowl.Hide
End Sub

Private Sub cmdposition_Click()
'switch between forms
frmposition.Show
frmsuperbowl.Hide
End Sub

Private Sub cmdquit_Click()
'quit program
End
End Sub

Private Sub cmdsuperbowl_Click()
'identifying variables
Dim teamwinner(1 To 21) As String
Dim teamloser(1 To 21) As String
Dim score(1 To 21) As String
Dim year(1 To 21) As Single
Dim location(1 To 21) As String
Dim bowl(1 To 21) As String
Dim P As Integer
Dim superbowlyear As Single

'setting what found equals to
found = False

'clear picture box
picresults.Cls
piclocation.Cls
picnewyear.Cls
picbowl.Cls
picresults2.Cls



'opens/access the data file
Open App.Path & "\champions.txt" For Input As #1

'filling an array from the specific data file
Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, year(ctr), teamwinner(ctr), teamloser(ctr), score(ctr), location(ctr), bowl(ctr)
Loop

'set variable equal to
superbowlyear = InputBox("Enter Superbowl Year")

'searches for the data until found or at the end of the list
Do While ((Not found) And (P < ctr))
    P = P + 1
    If superbowlyear = year(P) Then
        found = True
    End If
Loop

'print results depending on variable
If (Not found) Then
        picresults.Print "Information is Not Available"
    Else
        picresults.Print ; teamwinner(P);
        picresults2.Print ; teamloser(P);
        piclocation.Print ; location(P);
        picnewyear.Print ; year(P);
        picbowl.Print ; bowl(P);
End If

'print results depending on variable value
Select Case superbowlyear
    Case Is = 1990
        picteams.Picture = LoadPicture(App.Path & "\49ers.JPG")
        picteams2.Picture = LoadPicture(App.Path & "\broncos.JPG")
    Case 1991
        picteams.Picture = LoadPicture(App.Path & "\giants.JPG")
        picteams2.Picture = LoadPicture(App.Path & "\bills.JPG")
    Case 1992
        picteams.Picture = LoadPicture(App.Path & "\redskins.JPG")
        picteams2.Picture = LoadPicture(App.Path & "\bills.JPG")
    Case 1993
        picteams.Picture = LoadPicture(App.Path & "\cowboys.JPG")
        picteams2.Picture = LoadPicture(App.Path & "\bills.JPG")
    Case 1994
        picteams.Picture = LoadPicture(App.Path & "\cowboys.JPG")
        picteams2.Picture = LoadPicture(App.Path & "\bills.JPG")
    Case 1995
        picteams.Picture = LoadPicture(App.Path & "\49ers.JPG")
        picteams2.Picture = LoadPicture(App.Path & "\chargers.JPG")
    Case 1996
        picteams.Picture = LoadPicture(App.Path & "\cowboys.JPG")
        picteams2.Picture = LoadPicture(App.Path & "\steelers.JPG")
    Case 1997
        picteams.Picture = LoadPicture(App.Path & "\packers.JPG")
        picteams2.Picture = LoadPicture(App.Path & "\patriots.JPG")
    Case 1998
        picteams.Picture = LoadPicture(App.Path & "\broncos.JPG")
        picteams2.Picture = LoadPicture(App.Path & "\packers.JPG")
    Case 1999
        picteams.Picture = LoadPicture(App.Path & "\broncos.JPG")
        picteams2.Picture = LoadPicture(App.Path & "\falcons.JPG")
    Case 2000
        picteams.Picture = LoadPicture(App.Path & "\rams.JPG")
        picteams2.Picture = LoadPicture(App.Path & "\titans.JPG")
    Case 2001
        picteams.Picture = LoadPicture(App.Path & "\ravens.JPG")
        picteams2.Picture = LoadPicture(App.Path & "\giants.JPG")
    Case 2002
        picteams.Picture = LoadPicture(App.Path & "\patriots.JPG")
        picteams2.Picture = LoadPicture(App.Path & "\rams.JPG")
    Case 2003
        picteams.Picture = LoadPicture(App.Path & "\buccaneers.JPG")
        picteams2.Picture = LoadPicture(App.Path & "\raiders.JPG")
    Case 2004
        picteams.Picture = LoadPicture(App.Path & "\patriots.JPG")
        picteams2.Picture = LoadPicture(App.Path & "\panthers.JPG")
    Case 2005
        picteams.Picture = LoadPicture(App.Path & "\patriots.JPG")
        picteams2.Picture = LoadPicture(App.Path & "\eagles.JPG")
    Case 2006
        picteams.Picture = LoadPicture(App.Path & "\steelers.JPG")
        picteams2.Picture = LoadPicture(App.Path & "\seahawks.JPG")
    Case 2007
        picteams.Picture = LoadPicture(App.Path & "\colts.JPG")
        picteams2.Picture = LoadPicture(App.Path & "\bears.JPG")
    Case 2008
        picteams.Picture = LoadPicture(App.Path & "\giants.JPG")
        picteams2.Picture = LoadPicture(App.Path & "\patriots.JPG")
    Case 2009
        picteams.Picture = LoadPicture(App.Path & "\steelers.JPG")
        picteams2.Picture = LoadPicture(App.Path & "\cardinals.JPG")
    Case 2010
        picteams.Picture = LoadPicture(App.Path & "\saints.JPG")
        picteams2.Picture = LoadPicture(App.Path & "\colts.JPG")
End Select

'closes data file to be re-read
Close

End Sub


Private Sub cmdwin_Click()
'identifying variables
Dim teamwinner(1 To 21) As String
Dim teamloser(1 To 21) As String
Dim score(1 To 21) As String
Dim year(1 To 21) As Single
Dim location(1 To 21) As String
Dim bowl(1 To 21) As String
Dim P As Integer
Dim superbowlyear As Single

'setting what found equals to
found = False

'clear picture box
picwin.Cls

'opens/access the data file
Open App.Path & "\champions.txt" For Input As #1

'filling an array from the specific data file
Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, year(ctr), teamwinner(ctr), teamloser(ctr), score(ctr), location(ctr), bowl(ctr)
Loop

'print results
picwin.Print ; teamwinner(1); , ; year(1)
picwin.Print ; teamwinner(2); , ; year(2)
picwin.Print ; teamwinner(3); , ; year(3)
picwin.Print ; teamwinner(4); , ; year(4)
picwin.Print ; teamwinner(5); , ; year(5)
picwin.Print ; teamwinner(6); , ; year(6)
picwin.Print ; teamwinner(7); , ; year(7)
picwin.Print ; teamwinner(8); , ; year(8)
picwin.Print ; teamwinner(9); , ; year(9)
picwin.Print ; teamwinner(10); , ; year(10)
picwin.Print ; teamwinner(11); , ; year(11)
picwin.Print ; teamwinner(12); , ; year(12)
picwin.Print ; teamwinner(13); , ; year(13)
picwin.Print ; teamwinner(14); , ; year(14)
picwin.Print ; teamwinner(15); , ; year(15)
picwin.Print ; teamwinner(16); , ; year(16)
picwin.Print ; teamwinner(17); , ; year(17)
picwin.Print ; teamwinner(18); , ; year(18)
picwin.Print ; teamwinner(19); , ; year(19)
picwin.Print ; teamwinner(20); , ; year(20)
picwin.Print ; teamwinner(21); , ; year(21)

'close data file to be re-open
Close

End Sub
