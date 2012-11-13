VERSION 5.00
Begin VB.Form LetsSeeTheCoaches 
   BackColor       =   &H0000FF00&
   Caption         =   "Form6"
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8985
   LinkTopic       =   "Form6"
   ScaleHeight     =   6375
   ScaleWidth      =   8985
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   615
      Left            =   7920
      TabIndex        =   9
      Top             =   4200
      Width           =   855
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return To Previous Page"
      Height          =   615
      Left            =   6120
      TabIndex        =   8
      Top             =   5400
      Width           =   2055
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   615
      Left            =   4320
      TabIndex        =   7
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton cmdSort 
      Caption         =   "Sort By Name"
      Height          =   1215
      Left            =   6120
      TabIndex        =   6
      Top             =   3960
      Width           =   1575
   End
   Begin VB.PictureBox Picture4 
      Height          =   2295
      Left            =   2640
      Picture         =   "Form6.frx":0000
      ScaleHeight     =   2235
      ScaleWidth      =   1515
      TabIndex        =   5
      Top             =   3720
      Width           =   1575
   End
   Begin VB.PictureBox Picture3 
      Height          =   3735
      Left            =   0
      Picture         =   "Form6.frx":207CA
      ScaleHeight     =   3675
      ScaleWidth      =   2595
      TabIndex        =   4
      Top             =   2280
      Width           =   2655
   End
   Begin VB.PictureBox Picture2 
      Height          =   3735
      Left            =   2280
      Picture         =   "Form6.frx":4137C
      ScaleHeight     =   3675
      ScaleWidth      =   1515
      TabIndex        =   3
      Top             =   0
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   3255
      Left            =   0
      Picture         =   "Form6.frx":632B6
      ScaleHeight     =   3195
      ScaleWidth      =   2355
      TabIndex        =   2
      Top             =   0
      Width           =   2415
   End
   Begin VB.PictureBox pbxResults 
      Height          =   3855
      Left            =   3840
      ScaleHeight     =   3795
      ScaleWidth      =   4995
      TabIndex        =   1
      Top             =   0
      Width           =   5055
   End
   Begin VB.CommandButton cmdCoaches 
      Caption         =   "See All Packers History Of Coaches"
      Height          =   1215
      Left            =   4320
      TabIndex        =   0
      Top             =   3960
      Width           =   1575
   End
End
Attribute VB_Name = "LetsSeeTheCoaches"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strName(1 To 13) As String 'say that strName is only from 1 to 13 in its array'
Dim i As Integer 'says that i can only be an integer'
Dim x(1 To 13) As String 'says that x can only be from 1 to 13 in its array'

Private Sub cmdClear_Click()
pbxResults.Cls
'this will clear whatever is in the picture box'
End Sub

Private Sub cmdCoaches_Click()
Open "M:\cs130\sjbenfante\Coaches\Coaches.txt" For Input As #1
'this opens and read a file named under the M drive'
For i = 1 To 13 'i is equal only from 1 to 13'
    Input #1, strName(i), x(i) 'this will read the name and the years from the file that was opened and read'
    pbxResults.Print strName(i), Tab(30), x(i)
    'this prints out the names and years onto a picture box'
Next i 'says to go to the next i indicated above'
Close #1 'this closes the file that you opened and read from'
End Sub


Private Sub cmdQuit_Click()
    End
'this will automatically end the program'
End Sub

Private Sub cmdReturn_Click()
LetsSeeTheCoaches.Hide
WhoCoachedWhen.Show
'this will hide the sixth form and show the fourth form'
End Sub

Private Sub cmdSort_Click()
Dim Pass As Integer 'says that the passes can only be integers'
Dim i As Integer 'say that i can only be an integer'
Dim temp As String 'say that temp can only be words or letters no integers'
Dim temp2 As String 'say that temp2 can only be words or letters no integers'
pbxResults.Cls 'this will clear what is in the picture box before the user does another action'
For Pass = 1 To 13 'only pass from 1 to 13'
    For i = 1 To 13 - Pass 'this takes away a pass as you go through it'
        If strName(i) < strName(i + 1) Then 'says to look at the first two names and if ___ then ___'
            temp = strName(i) 'puts the first name into a holder called temp'
            strName(i) = strName(i + 1) 'puts the first name into the second names spot'
            strName(i + 1) = temp 'puts the second name into the first names spot'
            temp2 = x(i) 'puts the first years into a holder called temp2'
            x(i) = x(i + 1) 'puts the first years into the second years spot'
            x(i + 1) = temp2 'puts the second years into the first years spot'
        End If 'this ends the If Then statement'
    Next i 'goes to the next i in the loop'
    pbxResults.Print strName(i), Tab(30), x(i)
    'this prints the name and the years in alphabetical order onto the picture box'
Next Pass 'this say to go to the next name'

End Sub

Private Sub Form_Load()
strPath = "n:\CS130\handin\sjbenfante\"
End Sub


