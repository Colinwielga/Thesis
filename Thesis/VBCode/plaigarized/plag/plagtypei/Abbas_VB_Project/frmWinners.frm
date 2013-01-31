VERSION 5.00
Begin VB.Form frmWinners
   BackColor       =   &H80000006&
   Caption         =   "Form1"
   ClientHeight    =   11745
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13515
   BeginProperty Font
      Name            =   "Courier New"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   11745
   ScaleWidth      =   13515
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit
      BackColor       =   &H0000C0C0&
      Caption         =   "Exit the Program"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   10080
      Width           =   2655
   End
   Begin VB.CommandButton cmdMainMenu
      BackColor       =   &H0000C0C0&
      Caption         =   "Go Back to the Main Menu"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   10080
      Width           =   2655
   End
   Begin VB.CommandButton cmdDisplay
      BackColor       =   &H0000C0C0&
      Caption         =   "Display Winners By Year"
      Enabled         =   0   'False
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1560
      Width           =   3015
   End
   Begin VB.CommandButton cmdSearchPosition
      BackColor       =   &H0000C0C0&
      Caption         =   "Search Winners By Position"
      Enabled         =   0   'False
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8760
      Width           =   3015
   End
   Begin VB.CommandButton cmdSearchPlayer
      BackColor       =   &H0000C0C0&
      Caption         =   "Search For Your Favorite Player"
      Enabled         =   0   'False
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7560
      Width           =   3015
   End
   Begin VB.CommandButton cmdSearchTeam
      BackColor       =   &H0000C0C0&
      Caption         =   "Search For Your Favorite School"
      Enabled         =   0   'False
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6360
      Width           =   3015
   End
   Begin VB.CommandButton cmdSortTeam
      BackColor       =   &H0000C0C0&
      Caption         =   "Sort Winners by School"
      Enabled         =   0   'False
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5160
      Width           =   3015
   End
   Begin VB.CommandButton cmdSortPosition
      BackColor       =   &H0000C0C0&
      Caption         =   "Sort Winners by Position"
      Enabled         =   0   'False
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3960
      Width           =   3015
   End
   Begin VB.CommandButton cmdSortAlpha
      BackColor       =   &H0000C0C0&
      Caption         =   "Sort Winners Alphabetically"
      Enabled         =   0   'False
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2760
      Width           =   3015
   End
   Begin VB.CommandButton cmdLoad
      BackColor       =   &H0000C0C0&
      Caption         =   "Load the Heisman Winners"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   4815
   End
   Begin VB.PictureBox picResults
      BackColor       =   &H80000009&
      BeginProperty Font
         Name            =   "Lucida Console"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10815
      Left            =   6600
      ScaleHeight     =   10755
      ScaleWidth      =   6075
      TabIndex        =   0
      Top             =   360
      Width           =   6135
   End
End
Attribute VB_Name = "frmWinners"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'The Heisman Trophy
'frmWinners
'Kevin Abbas
'2-16-10
'Objective of form -  To display winners of the Heisman trophy and allow the user to sort and search the data.

' asdfsdaf sdf sdafasdf agr
Dim Year(1 To 1000) As Integer, FirstName(1 To 1000) As String, LastName(1 To 1000) As String, School(1 To 1000) As String, Position(1 To 1000) As String, Ctr As Integer, Temp As String, Pass As Integer, Pos As Integer, S As Integer

' asldkf jsdfsdjfosdijf sd
Private Sub cmdDisplay_Click() 'display the winners by year - as they are in the data file
    Dim N As Integer
    picResults.Cls
        picResults.Print "Year", "First Name", "Last Name"; Tab(45); "School"; Tab(65); "Position"
        picResults.Print "**************************************************************************************"
    For N = 1 To Ctr
        picResults.Print Year(N), FirstName(N), LastName(N); Tab(45); School(N); Tab(65); Position(N)
    Next N
End Sub

' awoief sofisdfoifsdofisf
Private Sub cmdExit_Click() 'thank the user and exit the program
    MsgBox ("Hope you enjoyed learning about the Heisman, have a nice day!")
    End
End Sub

' laskdj fji i oifj fi sdi
Private Sub cmdLoad_Click() 'load the data from the data file
    picResults.Cls
    Open App.Path & "\Data.txt" For Input As #1
    Do While Not EOF(1)
        Ctr = Ctr + 1
        Input #1, Year(Ctr), FirstName(Ctr), LastName(Ctr), School(Ctr), Position(Ctr)
    Loop
    Close #1
    cmdLoad.Enabled = False 'Disable the load button
    cmdSearchPlayer.Enabled = True 'Enable the rest of the buttons
    cmdDisplay.Enabled = True
    cmdSearchPosition.Enabled = True
    cmdSearchTeam.Enabled = True
    cmdSortAlpha.Enabled = True
    cmdSortPosition.Enabled = True
    cmdSortTeam.Enabled = True
    MsgBox ("The Data has been Entered!")
End Sub

' lsdif dsi fiflsf sdlisd dsi i
Private Sub cmdMainMenu_Click() 'bring the user back to the main menu
    frmWelcome.Show
    frmHistory.Hide
    frmWinners.Hide
    frmWhereNow.Hide
End Sub

' os difds fsdil fsdl dsids fid
Private Sub cmdSearchPlayer_Click() 'search by player (last name)
    Dim Z As String, K As Integer, Found2 As Boolean
        Z = InputBox("Please Enter the Last Name of Your Favorite Player")
        picResults.Cls
        picResults.Print "Year", "First Name", "Last Name"; Tab(45); "School"; Tab(65); "Position"
        picResults.Print "**************************************************************************************"
        Found2 = False
        For K = 1 To Ctr
            If LastName(K) = Z Then
                picResults.Print Year(K), FirstName(K), LastName(K); Tab(45); School(K); Tab(65); Position(K)
                Found2 = True
            End If
        Next K
        If Found2 = False Then
            MsgBox ("Sorry! " & Z & " never won the Heisman!")
        End If
End Sub

  'lsfd lid ssdl fisdf jsdfi sdfsif
Private Sub cmdSearchPosition_Click() 'search by position
    Dim Q As String, P As Integer, Found3 As Boolean
        Q = InputBox("Please Enter The Position You Wish To Search")
        picResults.Cls
        picResults.Print "Year", "First Name", "Last Name"; Tab(45); "School"; Tab(65); "Position"
        picResults.Print "**************************************************************************************"
        Found3 = False
        For P = 1 To Ctr
            If Position(P) = Q Then
                picResults.Print Year(P), FirstName(P), LastName(P); Tab(45); School(P); Tab(65); Position(P)
                Found3 = True
            End If
        Next P
        If Found3 = False Then
            MsgBox ("Sorry! No " & P & "'s has ever won the Heisman!")
        End If
End Sub

    ' slfi dfsd fdls fdlsif dfidsfsdi fsdfsdf n
Private Sub cmdSearchTeam_Click() 'search by school
        Dim X As String, N As Integer, Found As Boolean
        X = InputBox("Please Enter Your Favorite School")
        picResults.Cls
        picResults.Print "Year", "First Name", "Last Name"; Tab(45); "School"; Tab(65); "Position"
        picResults.Print "**************************************************************************************"
        Found = False
        For N = 1 To Ctr
            If School(N) = X Then
                picResults.Print Year(N), FirstName(N), LastName(N); Tab(45); School(N); Tab(65); Position(N)
                Found = True
            End If
        Next N
        If Found = False Then
            MsgBox ("Sorry! Noone from " & X & " has won the Heisman. Maybe next Year!")
        End If


End Sub

     ' ls dfisdl fsdilf hsdfsdf nsdf dsf dsf
Private Sub cmdSortAlpha_Click() 'sort winners alphabetically

        picResults.Cls
        picResults.Print "Year", "First Name", "Last Name"; Tab(45); "School"; Tab(65); "Position"
        picResults.Print "**************************************************************************************"
        For Pass = 1 To Ctr - 1
        For Pos = 1 To Ctr - Pass
            If LastName(Pos) > LastName(Pos + 1) Then
                Temp = LastName(Pos)
                LastName(Pos) = LastName(Pos + 1)
                LastName(Pos + 1) = Temp
                Temp = Year(Pos)
                Year(Pos) = Year(Pos + 1)
                Year(Pos + 1) = Temp
                Temp = FirstName(Pos)
                FirstName(Pos) = FirstName(Pos + 1)
                FirstName(Pos + 1) = Temp
                Temp = School(Pos)
                School(Pos) = School(Pos + 1)
                School(Pos + 1) = Temp
                Temp = Position(Pos)
                Position(Pos) = Position(Pos + 1)
                Position(Pos + 1) = Temp
            End If
        Next Pos
    Next Pass
    For S = 1 To Ctr
        picResults.Print Year(S), FirstName(S), LastName(S); Tab(45); School(S); Tab(65); Position(S)
    Next S
End Sub

 ' lsdfkjds flksdjflsdkf sdfsdlfsd ds
Private Sub cmdSortPosition_Click() 'sort winners by position

        picResults.Cls
        picResults.Print "Year", "First Name", "Last Name"; Tab(45); "School"; Tab(65); "Position"
        picResults.Print "**************************************************************************************"
        For Pass = 1 To Ctr - 1
        For Pos = 1 To Ctr - Pass
            If Position(Pos) > Position(Pos + 1) Then
                Temp = Position(Pos)
                Position(Pos) = Position(Pos + 1)
                Position(Pos + 1) = Temp
                Temp = Year(Pos)
                Year(Pos) = Year(Pos + 1)
                Year(Pos + 1) = Temp
                Temp = FirstName(Pos)
                FirstName(Pos) = FirstName(Pos + 1)
                FirstName(Pos + 1) = Temp
                Temp = School(Pos)
                School(Pos) = School(Pos + 1)
                School(Pos + 1) = Temp
                Temp = LastName(Pos)
                LastName(Pos) = LastName(Pos + 1)
                LastName(Pos + 1) = Temp
            End If
        Next Pos
    Next Pass
    For S = 1 To Ctr
        picResults.Print Year(S), FirstName(S), LastName(S); Tab(45); School(S); Tab(65); Position(S)
    Next S
End Sub

' lks dfjlsdf sdlf jsdilf sdifdsj
Private Sub cmdSortTeam_Click() 'sort winners by school

        picResults.Cls
        picResults.Print "Year", "First Name", "Last Name"; Tab(45); "School"; Tab(65); "Position"
        picResults.Print "**************************************************************************************"
        For Pass = 1 To Ctr - 1
        For Pos = 1 To Ctr - Pass
            If School(Pos) > School(Pos + 1) Then
                Temp = School(Pos)
                School(Pos) = School(Pos + 1)
                School(Pos + 1) = Temp
                Temp = Year(Pos)
                Year(Pos) = Year(Pos + 1)
                Year(Pos + 1) = Temp
                Temp = FirstName(Pos)
                FirstName(Pos) = FirstName(Pos + 1)
                FirstName(Pos + 1) = Temp
                Temp = Position(Pos)
                Position(Pos) = Position(Pos + 1)
                Position(Pos + 1) = Temp
                Temp = LastName(Pos)
                LastName(Pos) = LastName(Pos + 1)
                LastName(Pos + 1) = Temp
            End If
        Next Pos
    Next Pass
    For S = 1 To Ctr
        picResults.Print Year(S), FirstName(S), LastName(S); Tab(45); School(S); Tab(65); Position(S)
    Next S
End Sub

               ' sdlf dsfl dsfli sdfi ldifn dfndsfn dsf nds
    ' sdfld sflds flsdf s
Private Sub Form_Load()
    Top = Screen.Height / 2 - Height / 2
    Left = Screen.Width / 2 - Width / 2

End Sub
