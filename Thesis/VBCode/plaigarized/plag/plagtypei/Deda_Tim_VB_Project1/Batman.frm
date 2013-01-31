VERSION 5.00
Begin VB.Form Batman
   Caption         =   "Form1"
   ClientHeight    =   8880
   ClientLeft      =   3930
   ClientTop       =   3660
   ClientWidth     =   11805
   LinkTopic       =   "Form1"
   ScaleHeight     =   8880
   ScaleWidth      =   11805
   Begin VB.CommandButton Search
      BackColor       =   &H8000000C&
      Caption         =   "Search by Actor/Actress"
      Enabled         =   0   'False
      Height          =   735
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3480
      Width           =   1935
   End
   Begin VB.CommandButton Search2
      BackColor       =   &H8000000C&
      Caption         =   "Search by Character"
      Enabled         =   0   'False
      Height          =   735
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4560
      Width           =   1935
   End
   Begin VB.CommandButton BatmanQuiz
      BackColor       =   &H8000000C&
      Caption         =   "Continue to Batman Quiz"
      Height          =   615
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7680
      Width           =   1815
   End
   Begin VB.CommandButton TheDarkKnight
      BackColor       =   &H8000000C&
      Caption         =   "Print Cast for The Dark Knight"
      Height          =   735
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2400
      Width           =   1935
   End
   Begin VB.CommandButton BatmanBegins
      BackColor       =   &H8000000C&
      Caption         =   "Print Cast for Batman Begins"
      Height          =   735
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1320
      Width           =   1935
   End
   Begin VB.PictureBox PicResults
      BackColor       =   &H80000007&
      ForeColor       =   &H8000000E&
      Height          =   6255
      Left            =   720
      ScaleHeight     =   6195
      ScaleWidth      =   5715
      TabIndex        =   1
      Top             =   1080
      Width           =   5775
   End
   Begin VB.CommandButton Return
      BackColor       =   &H8000000C&
      Caption         =   "Return to Main Menu"
      Height          =   615
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7680
      Width           =   1815
   End
   Begin VB.Label Label1
      BackStyle       =   0  'Transparent
      Caption         =   "Batman Cast Lists"
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
      Left            =   720
      TabIndex        =   7
      Top             =   480
      Width           =   5055
   End
   Begin VB.Image Image1
      Height          =   18000
      Left            =   -1320
      Picture         =   "Batman.frx":0000
      Top             =   -3240
      Width           =   24000
   End
End
Attribute VB_Name = "Batman"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Declare form level variables
Dim BmanBeginsListAct(1 To 18) As String, DarkKnightListAct(1 To 28) As String
' inserted comment here
Dim BmanBeginsListChar(1 To 18) As String, DarkKnightListChar(1 To 28) As String
Dim CTR As Integer, CTR2 As Integer
' inserted comment here
Private Sub BatmanBegins_Click()
'initialize ctr to 1, to be used for position in the array
PicResults.Cls
' inserted comment here
CTR = 0

'Open and read file
Open App.Path & "\Batman Begins Movie.txt" For Input As #1

'print header
PicResults.Print "Names of Actor/Actress"; Tab(30); "Character"
        PicResults.Print "***************************************************************"
'set names to variables
' inserted comment here
    Do While Not EOF(1)
        'set counter and move to next number
        CTR = CTR + 1
        Input #1, BmanBeginsListAct(CTR), BmanBeginsListChar(CTR)
' inserted comment here
        'print names and cast
        PicResults.Print BmanBeginsListAct(CTR); Tab(30); BmanBeginsListChar(CTR)
    Loop
Close #1

'Enable search
Search.Enabled = True
' inserted comment here
Search2.Enabled = True
End Sub

Private Sub BatmanQuiz_Click()
' inserted comment here
Batman.Hide
Batman_Quiz.Show
End Sub


' inserted comment here
Private Sub Return_Click()
MainMenu.Show
Batman.Hide
' inserted comment here
End Sub

' inserted comment here
Private Sub Search_Click()
' inserted comment here
Dim Act As String, I As Integer, J As Integer, Found As Boolean, Found2 As Boolean
'clear picresults
    PicResults.Cls
'set counters adn variables
    I = 0
    J = 0
    Found = False
' inserted comment here
    Found2 = False
    Act = InputBox("Enter the name of the actor or actress you wish to find. Please put the actor's or actress's first and last name in that order with proper spacing. Also don't forget to capitalize as well!")
'search first movie
    Do While I < 18 And Found = False
        I = I + 1
        If Act = BmanBeginsListAct(I) Then
            Found = True
' inserted comment here
        End If
    Loop
'if still not found search second list
         Do While J < 28 And Found2 = False
            J = J + 1
' inserted comment here
            If Act = DarkKnightListAct(J) Then
                 Found2 = True
            End If
         Loop
'print
   If Found2 = True And Found = True Then
' inserted comment here
      PicResults.Print DarkKnightListAct(J); " was in Batman Begins and The Dark Knight"; Chr(13); " and played "; DarkKnightListChar(J)

    ElseIf Found = True Then
' inserted comment here
       PicResults.Print BmanBeginsListAct(I); " was in Batman Begins and played "; BmanBeginsListChar(I)
    ElseIf Found2 = True And Found = False Then
        PicResults.Print DarkKnightListAct(J); " was in The Dark Knight and played "; DarkKnightListChar(J)
    End If
'account for not found or both movies
If Found2 = False And Found = False Then
' inserted comment here
    MsgBox ("Sorry " & UserName & " The name you searched is either mispelled or does not exist.")
End If

End Sub

Private Sub Search2_Click()
' inserted comment here
Dim Char As String, K As Integer, L As Integer, Found As Boolean, Found2 As Boolean
'clear picresults
    PicResults.Cls
'set counters adn variables
    K = 0
    L = 0
' inserted comment here
    Found = False
    Found2 = False
    Char = InputBox("Enter the name of the character you wish to find. If the person has a secret identity the please put a slash and the masked ID second. I.E. Bruce Wayne/Batman. Don't forget to capitalize as well!")
'search first movie
    Do While K < 18 And Found = False
        K = K + 1
        If Char = BmanBeginsListChar(K) Then
            Found = True
        End If
    Loop

'if still not found search second list
         Do While L < 28 And Found2 = False
' inserted comment here
            L = L + 1
            If Char = DarkKnightListChar(L) Then
                Found2 = True
            End If
         Loop
'print
    If Found = True And Found2 = True Then
        PicResults.Print DarkKnightListChar(L); " was in Batman Begins and The Dark Knight"; Chr(13); " and was played by "; DarkKnightListAct(L)
    ElseIf Found = True Then
' inserted comment here
        PicResults.Print BmanBeginsListChar(K); " was in Batman Begins and was played by "; BmanBeginsListAct(K)
    ElseIf Found2 = True Then
        PicResults.Print DarkKnightListChar(L); " was in The Dark Knight and was played by "; DarkKnightListAct(L)
    End If
'account for not found
If Found2 = False And Found = False Then
    MsgBox ("Sorry " & UserName & " the character you searched is either mispelled or does not exist.")
End If
End Sub

Private Sub TheDarkKnight_Click()
' inserted comment here
'initialize ctr to 1, to be used for position in the array
PicResults.Cls
CTR2 = 0

'Open and read file
' inserted comment here
Open App.Path & "\The Dark Knight Movie.txt" For Input As #2

'print header
PicResults.Print "Names of Actor/Actress"; Tab(30); "Character"
        PicResults.Print "***************************************************************"
'set names to variables
    Do While Not EOF(2)
        'set counter and move to next number
' inserted comment here
        CTR2 = CTR2 + 1
        Input #2, DarkKnightListAct(CTR2), DarkKnightListChar(CTR2)
        'print names and cast
        PicResults.Print DarkKnightListAct(CTR2); Tab(30); DarkKnightListChar(CTR2)
    Loop
Close #2
'Enable search
Search.Enabled = True
' inserted comment here
Search2.Enabled = True
End Sub
