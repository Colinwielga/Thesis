VERSION 5.00
Begin VB.Form Ironman
   Caption         =   "Form1"
   ClientHeight    =   9120
   ClientLeft      =   3930
   ClientTop       =   3660
   ClientWidth     =   11835
   LinkTopic       =   "Form1"
   ScaleHeight     =   9120
   ScaleWidth      =   11835
   Begin VB.CommandButton Sort2
      BackColor       =   &H0000C0C0&
      Caption         =   "Sort by Character Name"
      Enabled         =   0   'False
      Height          =   735
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2280
      Width           =   1935
   End
   Begin VB.CommandButton Sort
      BackColor       =   &H0000C0C0&
      Caption         =   "Sort Cast List by Name of Actor/Actress"
      Enabled         =   0   'False
      Height          =   735
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1320
      Width           =   1935
   End
   Begin VB.CommandButton IronmanQuiz
      BackColor       =   &H0000C0C0&
      Caption         =   "Continue to Ironman Quiz"
      Height          =   735
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6840
      Width           =   1935
   End
   Begin VB.CommandButton CastList
      BackColor       =   &H0000C0C0&
      Caption         =   "Print the Cast List for Iron Man"
      Height          =   735
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1200
      Width           =   1935
   End
   Begin VB.PictureBox PicResults
      BackColor       =   &H000000C0&
      ForeColor       =   &H0000FFFF&
      Height          =   6855
      Left            =   2520
      ScaleHeight     =   6795
      ScaleWidth      =   4755
      TabIndex        =   1
      Top             =   1200
      Width           =   4815
   End
   Begin VB.CommandButton Return
      BackColor       =   &H0000C0C0&
      Caption         =   "Return to Main Menu"
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6840
      Width           =   2055
   End
   Begin VB.Label Label2
      BackStyle       =   0  'Transparent
      Caption         =   "Iron Man Cast Lists"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   2880
      TabIndex        =   7
      Top             =   480
      Width           =   4215
   End
   Begin VB.Label Label1
      BackStyle       =   0  'Transparent
      Caption         =   "Will sort into Alphabetical order"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   2175
   End
   Begin VB.Image Image1
      Height          =   15360
      Left            =   -4320
      Picture         =   "Ironman.frx":0000
      Top             =   -480
      Width           =   19200
   End
End
Attribute VB_Name = "Ironman"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CTR As Integer

Dim IronListAct(1 To 14) As String, IronListChar(1 To 14) As String

Private Sub CastList_Click()

'initialize ctr to 1, to be used for position in the array

PicResults.Cls

CTR = 0


'Open and read file
Open App.Path & "\Iron Man.txt" For Input As #1

'print header
PicResults.Print "Names of Actor/Actress"; Tab(30); "Character"
        PicResults.Print "***************************************************************"
'set names to variables
    Do While Not EOF(1)
        'set counter and move to next number
        CTR = CTR + 1
        Input #1, IronListAct(CTR), IronListChar(CTR)
        'print names and cast
        PicResults.Print IronListAct(CTR); Tab(30); IronListChar(CTR)
    Loop
Close #1

'Enable search

Sort.Enabled = True

Sort2.Enabled = True

End Sub


Private Sub IronmanQuiz_Click()

Ironman_Quiz.Show

Ironman.Hide

End Sub


Private Sub Return_Click()

MainMenu.Show

Ironman.Hide

End Sub


Private Sub Sort_Click()
'clear PicResults
PicResults.Cls
'dim variables
Dim pass As Integer, pos As Integer, I As Integer
Dim tempIronListAct As String, tempIronListChar As String

'sort the names
For pass = 1 To CTR - 1
    For pos = 1 To CTR - pass
        If IronListAct(pos) > IronListAct(pos + 1) Then
            tempIronListAct = IronListAct(pos)
            IronListAct(pos) = IronListAct(pos + 1)
            IronListAct(pos + 1) = tempIronListAct
            tempIronListChar = IronListChar(pos)
            IronListChar(pos) = IronListChar(pos + 1)
            IronListChar(pos + 1) = tempIronListChar
        End If
    Next pos
Next pass

'print header
PicResults.Print "Names of Actor/Actress"; Tab(30); "Character"
    PicResults.Print "***************************************************************"

'print info sorted
  For I = 1 To CTR
             PicResults.Print IronListAct(I); Tab(30); IronListChar(I)
    Next I
End Sub

Private Sub Sort2_Click()
'clear PicResults
PicResults.Cls
'dim variables
Dim pass As Integer, pos As Integer, I As Integer
Dim tempIronListAct As String, tempIronListChar As String

'sort the names

For pass = 1 To CTR - 1

    For pos = 1 To CTR - pass

        If IronListChar(pos) > IronListChar(pos + 1) Then

            tempIronListChar = IronListChar(pos)

            IronListChar(pos) = IronListChar(pos + 1)

            IronListChar(pos + 1) = tempIronListChar

            tempIronListAct = IronListAct(pos)

            IronListAct(pos) = IronListAct(pos + 1)

            IronListAct(pos + 1) = tempIronListAct

        End If

    Next pos

Next pass


'print header
PicResults.Print "Character"; Tab(30); "Names of Actor/Actress"
    PicResults.Print "***************************************************************"

'print info sorted

  For I = 1 To CTR

             PicResults.Print IronListChar(I); Tab(30); IronListAct(I)

    Next I

End Sub
