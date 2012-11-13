VERSION 5.00
Begin VB.Form frmDraft 
   BackColor       =   &H00FF0000&
   Caption         =   "2006 St. Johns Housing Draft"
   ClientHeight    =   7080
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10515
   LinkTopic       =   "Form1"
   ScaleHeight     =   7080
   ScaleWidth      =   10515
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOutput 
      BackColor       =   &H000000FF&
      Caption         =   "Output Housing Assignments To Text File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton cmdShow 
      BackColor       =   &H000000FF&
      Caption         =   "Display Housing Assignments"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5040
      Width           =   1935
   End
   Begin VB.CommandButton cmdDisplay 
      BackColor       =   &H000000FF&
      Caption         =   "Display Order For Draft"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2760
      Width           =   1935
   End
   Begin VB.PictureBox picDisplay 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5775
      Left            =   2880
      ScaleHeight     =   5715
      ScaleWidth      =   5595
      TabIndex        =   3
      Top             =   360
      Width           =   5655
   End
   Begin VB.CommandButton cmdLoad 
      BackColor       =   &H000000FF&
      Caption         =   "Load File Of Applicants"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   480
      Width           =   1935
   End
   Begin VB.CommandButton cmdDraft 
      BackColor       =   &H000000FF&
      Caption         =   "Continue With Draft"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3960
      Width           =   1935
   End
   Begin VB.CommandButton cmdSort 
      BackColor       =   &H000000FF&
      Caption         =   "Sort Housing Applicants"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1560
      UseMaskColor    =   -1  'True
      Width           =   1935
   End
   Begin VB.Label lblAuthor 
      BackColor       =   &H00FF0000&
      Caption         =   "Project Created By:                 Kyle Johnson"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8520
      TabIndex        =   7
      Top             =   5880
      Width           =   2055
   End
End
Attribute VB_Name = "frmDraft"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'St. Johns Housing Project
' Draft Form
' Written By Kyle Johnson
' 3/22/06
' The purpose of this form is to initialize the draft process by
'loading the file, sort the file, and begin the draft
' it also serves as a transition point between each student drafter
' finally, the form allows the output of the recieved data to a text file




Option Explicit
    Dim size As Integer

Private Sub cmdDisplay_Click()
    'Display the complete list of housing applicants along with thier position, and number of credits
    
    Dim J As Integer
        picDisplay.Cls
        picDisplay.Print "#"; Tab(14); "Name"; Tab(40); "Credits"; Tab(60)
        
        
        For J = 1 To size
            picDisplay.Print
            picDisplay.Print J; Tab(14); namesArray(J); Tab(40); creditArray(J); Tab(60);
            
            
          Next J
         'message box informs the user of how to proceed with the program
         
    MsgBox "Now You are Ready To Begin the Draft", , "Step 4"
End Sub


Private Sub cmdDraft_Click()
    'this takes the user to the draft page in order to pick the their housing.
    'once K exceeds the number of names in the array, then the draft is over, and it reports that all applicants have drafted
    
    K = K + 1
    
    If K <= size Then
        frmDraft.Visible = False
        frmOptions.Visible = True
        MsgBox namesArray(K) & " It is your turn to draft", , "Please Pick A House"
    Else
        MsgBox "All Students Have Been Assigned Housing", , "Draft Complete"
        
    End If
    
    'once the draft button is initially clicked, the load and sort buttons are hidden so that the user can not reload the file.
    
    cmdLoad.Visible = False
    cmdSort.Visible = False

End Sub

Private Sub cmdLoad_Click()

    'opens the appropriate text file, and saves the information in parrallel arrays
    
    pos = 0
        Open App.Path & "\credit.txt" For Input As #1
        
        Do Until EOF(1)
            pos = pos + 1
                Input #1, namesArray(pos), creditArray(pos)
            Loop
    Close #1
    
    'message box telling the user how to procede
    
    MsgBox "Now Sort All of the Housing Applicants by Credit", , " Step 2"
    'save the pos counter as size so it can be used by other sub-routines
    size = pos
End Sub

Private Sub cmdOutput_Click()
    'dim local variables
    Dim L As Integer
    'open a text file to be written in
    Open App.Path & "\housing.txt" For Output As #2
    ' for all locations in the array, print the name along with thier house to the file
    For L = 1 To K
        Print #2, namesArray(L); Tab(30); houseArray(L)
    Next L
    'close the file so it can be accessed by other programs
    Close #2
End Sub

Private Sub cmdShow_Click()
    'display the list of students along with the house they are assigned to for those that have already drafted
    Dim A As Integer
        picDisplay.Cls
        picDisplay.Print "Name"; Tab(30); "Housing Assignment"
        picDisplay.Print "**********************************************************************"
    
    For A = 1 To size
        picDisplay.Print namesArray(A); Tab(30); houseArray(A)
        picDisplay.Print
    Next A

End Sub

Private Sub cmdSort_Click()
    'sort the information in the array in descending order based on the credits
    'also sort the parrallel name array so the name stays with the correct number of credits
    
    Dim I, pass As Integer
    Dim tempCredit As Integer
    Dim tempName As String
    
    For pass = 1 To size - 1
        For I = 1 To size - pass
            If creditArray(I) < creditArray(I + 1) Then
                tempCredit = creditArray(I)
                creditArray(I) = creditArray(I + 1)
                creditArray(I + 1) = tempCredit
                tempName = namesArray(I)
                namesArray(I) = namesArray(I + 1)
                namesArray(I + 1) = tempName
            End If
        Next I
    Next pass
    'message box telling user how to procede
    
    MsgBox "Now Display The Order for the Draft", , " Step 3"

End Sub


