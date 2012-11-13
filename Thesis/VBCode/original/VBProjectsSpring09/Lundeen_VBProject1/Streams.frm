VERSION 5.00
Begin VB.Form Streams 
   BackColor       =   &H00FF80FF&
   Caption         =   "Form1"
   ClientHeight    =   8160
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9990
   LinkTopic       =   "Form1"
   ScaleHeight     =   8160
   ScaleWidth      =   9990
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Quit4 
      BackColor       =   &H00FF8080&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6360
      Width           =   2535
   End
   Begin VB.CommandButton RainbowSort 
      BackColor       =   &H00FF8080&
      Caption         =   "Let's Search For Streams With Rainbow Trout"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5040
      Width           =   2535
   End
   Begin VB.CommandButton BrownSort 
      BackColor       =   &H00FF8080&
      Caption         =   "Let's Search For Streams With Brown Trout"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3600
      Width           =   2535
   End
   Begin VB.CommandButton BrookSort 
      BackColor       =   &H00FF8080&
      Caption         =   "Let's Search The Streams For Brook Trout"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2160
      Width           =   2535
   End
   Begin VB.CommandButton Streams 
      BackColor       =   &H00FF8080&
      Caption         =   "Streams"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   600
      Width           =   2535
   End
   Begin VB.PictureBox Picture3 
      Height          =   6615
      Left            =   720
      ScaleHeight     =   6555
      ScaleWidth      =   4995
      TabIndex        =   0
      Top             =   480
      Width           =   5055
   End
End
Attribute VB_Name = "Streams"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This form gives the user some streams in Minnesota where trout can be found.  It then sorts them
'depending on what the user wants to fish for.
'Kevin Lundeen
'March 23

Option Explicit

'Declare the variables

Dim CTR As Integer, Stream(1 To 50) As String, Brook(1 To 50) As String, Brown(1 To 50) As String, Rainbow(1 To 50) As String

'This subroutine sorts the list depending on whether or not the stream has brook trout

Private Sub BrookSort_Click()
Dim J As Integer        'Declare a counter variable

    Picture3.Cls
    Picture3.Print "Stream", "", "Brook?", "Brown?", "Rainbow?"
    Picture3.Print "*****************************************************************************"
    
    J = 0                               'Set it equal to 0
    For J = 1 To CTR
        If Brook(J) = "Yes" Then        'This step determines if a stream has brook trout
            Picture3.Print Stream(J), "", Brook(J), Brown(J), Rainbow(J)
        End If
    Next J
        
End Sub

'This subroutine determines whether or not the list of streams has brown trout in it

Private Sub BrownSort_Click()
Dim K As Integer        'Declare a counter variable

    Picture3.Cls
    Picture3.Print "Stream", "", "Brook?", "Brown?", "Rainbow?"
    Picture3.Print "******************************************************************************"
    
    K = 0                               'Set variable equal to 0
    For K = 1 To CTR
        If Brown(K) = "Yes" Then        'This step determines if a stream has brown trout in it
            Picture3.Print Stream(K), , Brook(K), Brown(K), Rainbow(K)
        End If
    Next K
    
End Sub

Private Sub Quit4_Click()
    'This subroutine ends the program and thanks the user for going through it
    
    MsgBox ("Alright, so now you know the basics of what to use, what to fish for, and where to go.  Good Luck fishing!")
    End     'Ends program
End Sub

'This subroutine sorts the list of streams by whether or not they have rainbow trout

Private Sub RainbowSort_Click()
    Dim L As Integer        'Declare the variable
    
    
    Picture3.Cls
    Picture3.Print "Stream", "", "Brook?", "Brown?", "Rainbow?"
    Picture3.Print "***********************************************************************************"
    L = 0                                       'Set it equal to 0
    For L = 1 To CTR
        If Rainbow(L) = "Yes" Then              'This step determines if a stream does have rainbow trout in it
            Picture3.Print Stream(L), , Brook(L), Brown(L), Rainbow(L)
        End If
    Next L
End Sub

'This subroutine takes information from a data file and prints it in a picture box

Private Sub Streams_Click()


Open App.Path & "/TroutStreams.txt" For Input As #1                     'Opens the data file

    Picture3.Print "Stream", "", "Brook?", "Brown?", "Rainbow?"
    Picture3.Print "****************************************************************************************"

CTR = 0
Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, Stream(CTR), Brook(CTR), Brown(CTR), Rainbow(CTR)         'Reads the information
    Picture3.Print Stream(CTR), , Brook(CTR), Brown(CTR), Rainbow(CTR)  'Prints the information
Loop

End Sub




