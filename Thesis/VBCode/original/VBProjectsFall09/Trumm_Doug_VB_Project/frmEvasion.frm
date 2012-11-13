VERSION 5.00
Begin VB.Form frmEvasion 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Evading taxes"
   ClientHeight    =   11355
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   18720
   LinkTopic       =   "Form1"
   ScaleHeight     =   11355
   ScaleWidth      =   18720
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdSwitch 
      Caption         =   "Retirement Plan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   15120
      TabIndex        =   10
      Top             =   10200
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   11880
      TabIndex        =   9
      Top             =   10200
      Width           =   2655
   End
   Begin VB.CommandButton cmdDisplay 
      Caption         =   "Display the doctored data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   15120
      TabIndex        =   8
      Top             =   7920
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox txtReplacement 
      Height          =   975
      Left            =   8280
      TabIndex        =   6
      Top             =   9240
      Width           =   2655
   End
   Begin VB.TextBox txtNumberToReplace 
      Height          =   975
      Left            =   8280
      TabIndex        =   4
      Top             =   7920
      Width           =   2655
   End
   Begin VB.PictureBox Picture2 
      Height          =   5295
      Left            =   8760
      Picture         =   "frmEvasion.frx":0000
      ScaleHeight     =   5235
      ScaleWidth      =   8955
      TabIndex        =   3
      Top             =   1920
      Width           =   9015
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "Show actual importation data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   600
      TabIndex        =   2
      Top             =   7680
      Width           =   2655
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit data to avoid taxes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   11880
      TabIndex        =   1
      Top             =   7920
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.PictureBox picResults 
      Height          =   5415
      Left            =   600
      ScaleHeight     =   5355
      ScaleWidth      =   5715
      TabIndex        =   0
      Top             =   1800
      Width           =   5775
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "board room"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   1815
      Left            =   4680
      TabIndex        =   11
      Top             =   120
      Width           =   8775
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Enter a smaller replacement number to save on exportation fees===>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4320
      TabIndex        =   7
      Top             =   9120
      Width           =   3735
   End
   Begin VB.Label lblTextInput 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Enter the monthly bushel value to be replaced ===>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4320
      TabIndex        =   5
      Top             =   7920
      Width           =   3735
   End
End
Attribute VB_Name = "frmEvasion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    'This program works with the arrays used earlier
    'The Open For Output function is used to edit the file
    
    'Declare universal variables
    Dim bushelExports(1 To 20) As Integer, exportMonth(1 To 20), bCTR As Integer


Private Sub cmdShow_Click()
    'Declare varibles
    Dim J As Integer
    
    'Print header
    picResults.Print "Month"; Tab(20); "Thousands of bushels exported"
    picResults.Print "******************************************************************************************"
    
    'Prepare the file to be opened
    Open App.Path & "\bananasExports.txt" For Input As #1
    
    bCTR = 0
    
    Do While Not EOF(1)
        bCTR = bCTR + 1
        Input #1, bushelExports(bCTR), exportMonth(bCTR)
    Loop
    
    Close #1
    
    'For/Next Loop to print array data
    For J = 1 To bCTR
        picResults.Print exportMonth(J); Tab(20); bushelExports(J)
    Next J
    
    cmdEdit.Visible = True
End Sub

Private Sub cmdEdit_Click()
    Dim NumberToReplace As Integer, Replacement As Single, FoundTheValue As Boolean, Pos As Integer
    
    NumberToReplace = txtNumberToReplace.Text
    Replacement = txtReplacement.Text
    'Prepare the file to be opened and edited
    Open App.Path & "\bananasExports.txt" For Output As #2
    
    'Use Boolean to determine if value has been found
    FoundTheValue = False
    
    'For/Next loop search for a match, write feature outputs data in file
    For Pos = 1 To bCTR
        If (NumberToReplace <> bushels(Pos)) Then
            Write #2, bushelExports(Pos), exportMonth(Pos)
        Else
            FoundTheValue = True
            Write #2, Replacement, exportMonth(Pos)
        End If
    Next Pos
    
    'Message boxes tell user the outcome
    If FoundTheValue = True Then
        MsgBox ("Success! The value you identified was replaced with the number you chose.")
    Else
        MsgBox ("The number you wanted to remove was not in the data.  Try again.  Choose a bushel value to be removed.")
    End If
    
    Close #2
    
    cmdDisplay.Visible = True
    cmdShow.Visible = False
End Sub

Private Sub cmdDisplay_Click()
    'Declare variables
    Dim J As Integer

    picResults.Cls
    
    'Print header
    picResults.Print "Month"; Tab(20); "Thousands of bushels exported"
    picResults.Print "******************************************************************************************"
    
    'Prepare the file to be opened
    Open App.Path & "\bananasExports.txt" For Input As #1
    
    bCTR = 0
    
    'Reopen the file in order to show the change(s) the output function made
    Do While Not EOF(1)
        bCTR = bCTR + 1
        Input #1, bushelExports(bCTR), exportMonth(bCTR)
    Loop
    
    Close #1
    
    'For/Next Loop to print array data
    For J = 1 To bCTR
        picResults.Print exportMonth(J); Tab(20); bushelExports(J)
    Next J
    
    cmdSwitch.Visible = True
End Sub

Private Sub cmdStop_Click()
    End
End Sub

Private Sub cmdSwitch_Click()
    'Move on to next form
    frmEvasion.Hide
    frmRetire.Show
End Sub
