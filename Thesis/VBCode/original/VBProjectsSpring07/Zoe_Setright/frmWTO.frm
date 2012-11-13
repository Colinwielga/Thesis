VERSION 5.00
Begin VB.Form frmWTO 
   BackColor       =   &H00C0C0FF&
   Caption         =   "Form1"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12150
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   12150
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults2 
      BackColor       =   &H00FFFFC0&
      Height          =   10095
      Left            =   480
      ScaleHeight     =   10035
      ScaleWidth      =   7275
      TabIndex        =   5
      Top             =   480
      Width           =   7335
   End
   Begin VB.CommandButton cmdSortMembers 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Sort Members"
      Height          =   1575
      Left            =   8280
      TabIndex        =   4
      Top             =   6720
      Width           =   3255
   End
   Begin VB.CommandButton cmdMonth 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Find Members based on Month  they entered the WTO"
      Height          =   1575
      Left            =   8280
      TabIndex        =   3
      Top             =   4560
      Width           =   3375
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Quit"
      Height          =   1455
      Left            =   8280
      TabIndex        =   2
      Top             =   9600
      Width           =   3375
   End
   Begin VB.CommandButton cmdFindObserver 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Find Observer"
      Height          =   1575
      Left            =   8280
      TabIndex        =   1
      Top             =   2520
      Width           =   3255
   End
   Begin VB.CommandButton cmdFindMember 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Find Member"
      Height          =   1575
      Left            =   8280
      TabIndex        =   0
      Top             =   480
      Width           =   3255
   End
End
Attribute VB_Name = "frmWTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdFindMember_Click()
    Dim WTO(1 To 300) As String
    Dim Month(1 To 300) As String
    Dim Year(1 To 300) As Single
    Dim Ctr As Integer
    Dim Pos As Integer
    Dim Member As String
    Dim Found As Boolean
    
    
    Open App.Path & "\wto.txt" For Input As #1  'Retrieves the file to read
    Ctr = 0
    
    Do Until EOF(1)
        Ctr = Ctr + 1
        Input #1, WTO(Ctr), Month(Ctr), Year(Ctr)   'sorts the files into 3 seperate arrays
    Loop
    Close #1
    
    Member = InputBox("Enter in the member name you wish to find.", "Member")   'Gets the name from the user to search the array with.
    Pos = 0
    Found = False
    
    Do While ((Not Found) And (Pos < Ctr))  'This will search through the array in order to determine if there is a match.
        Pos = Pos + 1
        If Member = WTO(Pos) Then Found = True
    Loop
    
    
    If (Not Found) Then 'If no match was found a message will appear to the user
            MsgBox "This member was not found please see if they are an observer.", , "Not Found"
        Else    'This will clear the previous match and print the most recent
            picResults2.Cls
            picResults2.Print WTO(Pos), Tab(30); Month(Pos), Year(Pos)
    End If
    
        

End Sub

Private Sub cmdFindObserver_Click()
    Dim Observer(1 To 300) As String
    Dim Ctr As Integer
    Dim Pos As Integer
    Dim Found As Boolean
    Dim Ob As String
    
    Open App.Path & "\observer.txt" For Input As #2 'Retrieves the file to read
    Ctr = 0
    Do Until EOF(2)
        Ctr = Ctr + 1
        Input #2, Observer(Ctr) 'sorts the file into an array
    Loop
    Close #2
    
    Ob = InputBox("Enter a country name to see if they are an observer.", "Observer")   'Gets the name from the user to search the array with.
    Pos = 0
    Found = False
    
    Do While ((Not Found) And (Pos < Ctr))  'This will search through the array in order to determine if there is a match.
        Pos = Pos + 1
        If Ob = Observer(Pos) Then Found = True
    Loop
    
    If (Not Found) Then 'If no match was found a message will appear to the user
            picResults2.Print Ob; "is not an observer of the WTO."
        Else    'This will clear the previous match and print the most recent
            picResults2.Cls
            picResults2.Print Observer(Pos)
    End If
End Sub


Private Sub cmdMonth_Click()
    Dim WTO(1 To 300) As String
    Dim Month(1 To 300) As String
    Dim Year(1 To 300) As Single
    Dim M As String
    Dim Ctr As Integer
    Dim Pos As Integer
    Dim Found As Boolean
    Dim January As String
    
    Open App.Path & "\wto.txt" For Input As #3  'Retrieves the file to read
    Ctr = 0
    Do Until EOF(3)
        Ctr = Ctr + 1
        Input #3, WTO(Ctr), Month(Ctr), Year(Ctr)   'sorts the files into 3 seperate arrays
    Loop
    Close #3
    
    M = InputBox("Enter in a Month to search for members.", "Search by Entry into the WTO") 'Gets the name from the user to search the array with.
    Pos = 0
    Found = False
    picResults2.Print
    
    For Pos = 1 To Ctr  'This will go through the entire array and print every match
        If M = Month(Pos) Then
            picResults2.Print WTO(Pos), Tab(30); Month(Pos), Year(Pos)
        End If
    Next Pos

End Sub

Private Sub cmdQuit_Click()
    End
End Sub


Private Sub cmdSortMembers_Click()
    frmResults.Show
    frmWTO.Hide
End Sub
