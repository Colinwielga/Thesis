VERSION 5.00
Begin VB.Form frmOLC 
   BackColor       =   &H00004000&
   Caption         =   "OLC"
   ClientHeight    =   8760
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8850
   LinkTopic       =   "Form1"
   ScaleHeight     =   8760
   ScaleWidth      =   8850
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClimbing 
      Caption         =   "Climbing Wall"
      Height          =   495
      Left            =   960
      TabIndex        =   8
      Top             =   1200
      Width           =   1815
   End
   Begin VB.PictureBox pbxPrice 
      BackColor       =   &H00404040&
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   600
      ScaleHeight     =   915
      ScaleWidth      =   2595
      TabIndex        =   7
      Top             =   4320
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton cmdHours 
      Caption         =   "Business Hours"
      Height          =   495
      Left            =   960
      TabIndex        =   6
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CommandButton cmdPrice 
      Caption         =   "Rental Prices"
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   3720
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton CmdRentals 
      Caption         =   "Rental Items"
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   3120
      Width           =   1815
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   7320
      TabIndex        =   2
      Top             =   8280
      Width           =   1215
   End
   Begin VB.CommandButton cmdOLC 
      Caption         =   "What is the OLC?"
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Top             =   600
      Width           =   1815
   End
   Begin VB.PictureBox pbxResults 
      BackColor       =   &H00404040&
      ForeColor       =   &H00FFFFFF&
      Height          =   7455
      Left            =   4320
      ScaleHeight     =   7395
      ScaleWidth      =   4275
      TabIndex        =   0
      Top             =   480
      Width           =   4335
   End
   Begin VB.Label lblQuestions 
      BackColor       =   &H00004000&
      Caption         =   "If you have any questions please stop in and see us.  We are located in Mary Hall basement Room 033."
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   360
      TabIndex        =   10
      Top             =   8040
      Width           =   3975
   End
   Begin VB.Label lblWelcome 
      BackColor       =   &H00004000&
      Caption         =   "Welcome to the Outdoor Leadership Center"
      ForeColor       =   &H00FFFFFF&
      Height          =   350
      Left            =   3000
      TabIndex        =   9
      Top             =   120
      Width           =   3255
   End
   Begin VB.Image Image1 
      Height          =   2655
      Left            =   360
      Picture         =   "Project.frx":0000
      Top             =   5400
      Width           =   3435
   End
   Begin VB.Label lblRentals 
      BackColor       =   &H00004000&
      Caption         =   "The Outdoor Leadership Center offers a wide variety of things for students to rent.  Click below to see what we have to offer..."
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   360
      TabIndex        =   4
      Top             =   2400
      Width           =   3375
   End
End
Attribute VB_Name = "frmOLC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strItem(1 To 37) As String
Dim iId(1 To 37) As Integer
Dim iFirst(1 To 37) As Integer
Dim iAdd(1 To 37) As Integer
Dim strPath As String

Private Sub Form1_Load()
    strPath = "N:\CS130\KCKostka\"

End Sub


Private Sub cmdClimbing_Click()
    pbxResults.Cls
    Open strPath & "Climb.txt" For Input As #1
    Input #1, strClimb
        pbxResults.Print strClimb
    Close #1
    
    
End Sub

Private Sub cmdHours_Click()
cmdPrice.Visible = False
pbxPrice.Visible = False

    pbxResults.Cls
    Open strPath & "Hours.txt" For Input As #1
    Input #1, strHours
        pbxResults.Print strHours
    Close #1
End Sub

Private Sub cmdOLC_Click()
    cmdPrice.Visible = False
    pbxPrice.Visible = False
    
    pbxResults.Cls
    Open strPath & "Mission.txt" For Input As #1
    Input #1, Mission
        pbxResults.Print Mission    'Prints mission statement.
    Close #1
End Sub

Private Sub cmdPrice_Click()
    Open strPath & "PriceList.txt" For Input As #1 'Opens price list of rental equipment
    For I = 1 To 36
        Input #1, iId(I), strItem(I), iFirst(I), iAdd(I) 'stores data from document
    Next I
    Close #1
    
   Dim Found As Boolean
   Dim NewItem As Integer
   Dim K As Integer
        pbxPrice.Visible = True
        NewItem = InputBox("Enter item ID number", "Item Rental Price") '
        Found = False
        Do Until Found Or K = 36
            K = K + 1
            If NewItem = iId(K) Then
                Found = True
            End If
        Loop
        If Found = True Then
            pbxPrice.Cls
            pbxPrice.Print strItem(K)
            pbxPrice.Print "First day:"; (tab5); FormatCurrency(iFirst(K), 2)
            pbxPrice.Print "Additional Day(s):"; (tab5); FormatCurrency(iAdd(K), 2); " per day"
        Else
            MsgBox "No such item.  Enter correct ID number", , "Item not found"
        End If
End Sub

Private Sub cmdQuit_Click()
    End     'Exits program
End Sub

Private Sub CmdRentals_Click()
    cmdPrice.Visible = True
    pbxPrice.Visible = False
    Open strPath & "PriceList.txt" For Input As #1 'Opens price list of rental equipment
    For I = 1 To 36
        Input #1, iId(I), strItem(I), iFirst(I), iAdd(I) 'stores data from document
        
    Next I
    Close #1
    
    pbxResults.Cls 'clear picture box.
    For I = 1 To 36
        pbxResults.Print iId(I); (tab5); strItem(I)
    Next I
    
End Sub

