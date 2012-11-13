VERSION 5.00
Begin VB.Form frmGameBoard 
   BackColor       =   &H000080FF&
   Caption         =   "Game Board"
   ClientHeight    =   7920
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   FillColor       =   &H000080FF&
   ForeColor       =   &H000080FF&
   LinkTopic       =   "Form1"
   Picture         =   "frmGameBoard.frx":0000
   ScaleHeight     =   8.00000e5
   ScaleMode       =   0  'User
   ScaleWidth      =   8.00000e5
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLit200 
      BackColor       =   &H00FF0000&
      Caption         =   "$200"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   1320
      Width           =   2415
   End
   Begin VB.CommandButton cmdLit400 
      BackColor       =   &H00FF0000&
      Caption         =   "$400"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   2640
      UseMaskColor    =   -1  'True
      Width           =   2415
   End
   Begin VB.CommandButton cmdLit600 
      BackColor       =   &H00FF0000&
      Caption         =   "$600"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   3960
      UseMaskColor    =   -1  'True
      Width           =   2415
   End
   Begin VB.CommandButton cmdLit800 
      BackColor       =   &H00FF0000&
      Caption         =   "$800"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   5280
      Width           =   2415
   End
   Begin VB.CommandButton cmdLit1000 
      BackColor       =   &H00FF0000&
      Caption         =   "$1000"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   6600
      Width           =   2415
   End
   Begin VB.CommandButton cmdHis200 
      BackColor       =   &H00FF0000&
      Caption         =   "$200"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   1320
      Width           =   2415
   End
   Begin VB.CommandButton cmdHis400 
      BackColor       =   &H00FF0000&
      Caption         =   "$400"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   2640
      Width           =   2415
   End
   Begin VB.CommandButton cmdHis600 
      BackColor       =   &H00FF0000&
      Caption         =   "$600"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   3960
      Width           =   2415
   End
   Begin VB.CommandButton cmdHis800 
      BackColor       =   &H00FF0000&
      Caption         =   "$800"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   5280
      Width           =   2415
   End
   Begin VB.CommandButton cmdHis1000 
      BackColor       =   &H00FF0000&
      Caption         =   "$1000"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   6600
      Width           =   2415
   End
   Begin VB.CommandButton cmdMath200 
      BackColor       =   &H00FF0000&
      Caption         =   "$200"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   1320
      Width           =   2415
   End
   Begin VB.CommandButton cmdMath400 
      BackColor       =   &H00FF0000&
      Caption         =   "$400"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   2640
      Width           =   2415
   End
   Begin VB.CommandButton cmdMath600 
      BackColor       =   &H00FF0000&
      Caption         =   "$600"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   3960
      Width           =   2415
   End
   Begin VB.CommandButton cmdMath800 
      BackColor       =   &H00FF0000&
      Caption         =   "$800"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   5280
      Width           =   2415
   End
   Begin VB.CommandButton cmdMath1000 
      BackColor       =   &H00FF0000&
      Caption         =   "$1000"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   6600
      Width           =   2415
   End
   Begin VB.CommandButton cmdCartoons200 
      BackColor       =   &H00FF0000&
      Caption         =   "$200"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   1320
      Width           =   2415
   End
   Begin VB.CommandButton cmdCartoons400 
      BackColor       =   &H00FF0000&
      Caption         =   "$400"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   2640
      Width           =   2415
   End
   Begin VB.CommandButton cmdCartoons600 
      BackColor       =   &H00FF0000&
      Caption         =   "$600"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3960
      Width           =   2415
   End
   Begin VB.CommandButton cmdCartoons800 
      BackColor       =   &H00FF0000&
      Caption         =   "$800"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5280
      Width           =   2415
   End
   Begin VB.CommandButton cmdCartoons1000 
      BackColor       =   &H00FF0000&
      Caption         =   "$1000"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6600
      Width           =   2415
   End
   Begin VB.CommandButton cmdPlaces200 
      BackColor       =   &H00FF0000&
      Caption         =   "$200"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1320
      Width           =   2415
   End
   Begin VB.CommandButton cmdPlaces400 
      BackColor       =   &H00FF0000&
      Caption         =   "$400"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2640
      Width           =   2415
   End
   Begin VB.CommandButton cmdPlaces600 
      BackColor       =   &H00FF0000&
      Caption         =   "$600"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3960
      Width           =   2415
   End
   Begin VB.CommandButton cmdPlaces800 
      BackColor       =   &H00FF0000&
      Caption         =   "$800"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5280
      Width           =   2415
   End
   Begin VB.CommandButton cmdPlaces1000 
      BackColor       =   &H00FF0000&
      Caption         =   "$1000"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6600
      Width           =   2415
   End
   Begin VB.CommandButton cmdof200 
      BackColor       =   &H00FF0000&
      Caption         =   "$200"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   12720
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1320
      Width           =   2415
   End
   Begin VB.CommandButton cmdof400 
      BackColor       =   &H00FF0000&
      Caption         =   "$400"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   12720
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2640
      Width           =   2415
   End
   Begin VB.CommandButton cmdof600 
      BackColor       =   &H00FF0000&
      Caption         =   "$600"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   12720
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3960
      Width           =   2415
   End
   Begin VB.CommandButton cmdof800 
      BackColor       =   &H00FF0000&
      Caption         =   "$800"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   12720
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5280
      Width           =   2415
   End
   Begin VB.CommandButton cmdof1000 
      BackColor       =   &H00FF0000&
      Caption         =   "$1000"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   12720
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6600
      Width           =   2415
   End
   Begin VB.Label lblHistory 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "History"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   2640
      TabIndex        =   5
      Top             =   195
      Width           =   2415
   End
   Begin VB.Label lblDerivatives 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Derivatives"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   5160
      TabIndex        =   4
      Top             =   195
      Width           =   2415
   End
   Begin VB.Label lblCartoons 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Cartoon Characters"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   7680
      TabIndex        =   3
      Top             =   195
      Width           =   2415
   End
   Begin VB.Label lblPlaces 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   """P""laces on the Map"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   10200
      TabIndex        =   2
      Top             =   195
      Width           =   2415
   End
   Begin VB.Label lblof 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "_ of _"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   12720
      TabIndex        =   1
      Top             =   195
      Width           =   2415
   End
   Begin VB.Label lblLiterature 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Literature Authors"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   195
      Width           =   2415
   End
End
Attribute VB_Name = "frmGameBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This program will hide the command button pushed and hide the game board and show the forms corresponding to the command
'button pushed
'History for $200 is the daily double, so there are more instructions for that command button compared to the others!!!

Private Sub cmdCartoons1000_Click()
    
    'Shows and hides the forms
    frmCartoons1000.Show
    frmGameBoard.Hide
    
    'Hides the command button
    cmdCartoons1000.Visible = False
    
End Sub

Private Sub cmdCartoons200_Click()
    
    'Shows and hides the forms
    frmCartoons200.Show
    frmGameBoard.Hide
    
    'Hides the command button
    cmdCartoons200.Visible = False
    
End Sub

Private Sub cmdCartoons400_Click()
    
    'Shows and hides the forms
    frmCartoons400.Show
    frmGameBoard.Hide
    
    'Hides the command button
    cmdCartoons400.Visible = False
    
End Sub

Private Sub cmdCartoons600_Click()
    
    'Shows and hides the forms
    frmCartoons600.Show
    frmGameBoard.Hide
    
    'Hides the command button
    cmdCartoons600.Visible = False
    
End Sub

Private Sub cmdCartoons800_Click()
    
    'Shows and hides the forms
    frmCartoons800.Show
    frmGameBoard.Hide
    
    'Hides the command button
    cmdCartoons800.Visible = False
    
End Sub

Private Sub cmdHis1000_Click()
    
    'Shows and hides the forms
    frmHis1000.Show
    frmGameBoard.Hide
    
    'Hides the command button
    cmdHis1000.Visible = False
    
End Sub

Private Sub cmdHis200_Click()
    
    'Declaring variables
    frmDailyDouble.Show
    frmGameBoard.Hide
    cmdHis200.Visible = False
        
    'Setting initial value of wager
    Wager = 50000
        
    'Informing the user that they have found the daily double and asking them for their wager
    If Winnings >= 1000 Then
        Do Until Wager >= 0 And Wager <= Winnings And Wager / 100 = Int(Wager / 100)
            Wager = InputBox("Please enter the amount of your winnings that you would like to wager.  Your wager must be a multiple of 100 and must be positive!!!", "Amount of Wager")
            If Wager > Winnings Or Wager / 100 <> Int(Wager / 100) Or Wager < 0 Then
                MsgBox "You have entered an amount that is greater than your winnings or that is not a multiple of 100 or that is negative." & vbNewLine & "Please try again!", , "Cheater!!!"
            End If
        Loop
    Else
        MsgBox "Since your winnings is less than $1000, you can wager anything up to $1000", , "Winnings under $1000"
        Do Until Wager >= 0 And Wager <= 1000 And Wager / 100 = Int(Wager / 100)
            Wager = InputBox("Please enter your wager that must be a multiple of 100, less than or equal to 1000, and positive!!!", "Amount of Wager")
            If Wager > 1000 Or Wager / 100 <> Int(Wager / 100) Or Wager < 0 Then
                MsgBox "You have entered an amount that is greater than $1000 or that is not a multiple of 100 or that is negative." & vbNewLine & "Please try again!", , "Cheater!!!"
            End If
        Loop
    End If
    
    'Instructing the user what to do next
    MsgBox "Please click on the daily double to continue", , "Continue"
    
End Sub

Private Sub cmdHis400_Click()
    
    'Shows and hides the forms
    frmHis400.Show
    frmGameBoard.Hide
    
    'Hides the command button
    cmdHis400.Visible = False
    
End Sub

Private Sub cmdHis600_Click()
    
    'Shows and hides the forms
    frmHis600.Show
    frmGameBoard.Hide
    
    'Hides the command button
    cmdHis600.Visible = False
    
End Sub

Private Sub cmdHis800_Click()
    
    'Shows and hides the forms
    frmHis800.Show
    frmGameBoard.Hide
    
    'Hides the command button
    cmdHis800.Visible = False
    
End Sub

Private Sub cmdLit200_Click()
    
    'Shows and hides the forms
    frmLit200.Show
    frmGameBoard.Hide
    
    'Hides the command button
    cmdLit200.Visible = False
    
End Sub

Private Sub cmdLit400_Click()
    
    'Shows and hides the forms
    frmLit400.Show
    frmGameBoard.Hide
    
    'Hides the command button
    cmdLit400.Visible = False
    
End Sub

Private Sub cmdLit600_Click()
    'Shows and hides the forms
    frmLit600.Show
    frmGameBoard.Hide
    
    'Hides the command button
    cmdLit600.Visible = False
    
End Sub

Private Sub cmdLit800_Click()
    
    'Shows and hides the forms
    frmLit800.Show
    frmGameBoard.Hide
    
    'Hides the command button
    cmdLit800.Visible = False
    
End Sub

Private Sub cmdLit1000_Click()
    
    'Shows and hides the forms
    frmLit1000.Show
    frmGameBoard.Hide
    
    'Hides the command button
    cmdLit1000.Visible = False
    
End Sub

Private Sub cmdMath1000_Click()
    
    'Shows and hides the forms
    frmMath1000.Show
    frmGameBoard.Hide
    
    'Hides the command button
    cmdMath1000.Visible = False
    
End Sub

Private Sub cmdMath200_Click()
    
    'Shows and hides the forms
    frmMath200.Show
    frmGameBoard.Hide
    
    'Hides the command button
    cmdMath200.Visible = False
    
End Sub

Private Sub cmdMath400_Click()
    
    'Shows and hides the forms
    frmMath400.Show
    frmGameBoard.Hide
    
    'Hides the command button
    cmdMath400.Visible = False
    
End Sub

Private Sub cmdMath600_Click()
    
    'Shows and hides the forms
    frmMath600.Show
    frmGameBoard.Hide
    
    'Hides the command button
    cmdMath600.Visible = False
    
End Sub

Private Sub cmdMath800_Click()
    
    'Shows and hides the forms
    frmMath800.Show
    frmGameBoard.Hide
    
    'Hides the command button
    cmdMath800.Visible = False
    
End Sub

Private Sub cmdof1000_Click()
    
    'Shows and hides the forms
    frmof1000.Show
    frmGameBoard.Hide
    
    'Hides the command button
    cmdof1000.Visible = False
    
End Sub

Private Sub cmdof200_Click()
    
    'Shows and hides the forms
    frmof200.Show
    frmGameBoard.Hide
    
    'Hides the command button
    cmdof200.Visible = False
    
End Sub

Private Sub cmdof400_Click()
    
    'Shows and hides the forms
    frmof400.Show
    frmGameBoard.Hide
    
    'Hides the command button
    cmdof400.Visible = False
    
End Sub

Private Sub cmdof600_Click()
    
    'Shows and hides the forms
    frmof600.Show
    frmGameBoard.Hide
    
    'Hides the command button
    cmdof600.Visible = False
    
End Sub

Private Sub cmdof800_Click()
    
    'Shows and hides the forms
    frmof800.Show
    frmGameBoard.Hide
    
    'Hides the command button
    cmdof800.Visible = False
    
End Sub

Private Sub cmdPlaces1000_Click()
    
    'Shows and hides the forms
    frmPlaces1000.Show
    frmGameBoard.Hide
    
    'Hides the command button
    cmdPlaces1000.Visible = False
    
End Sub

Private Sub cmdPlaces200_Click()
    
    'Shows and hides the forms
    frmPlaces200.Show
    frmGameBoard.Hide
    
    'Hides the command button
    cmdPlaces200.Visible = False
    
End Sub

Private Sub cmdPlaces400_Click()
    
    'Shows and hides the forms
    frmPlaces400.Show
    frmGameBoard.Hide
    
    'Hides the command button
    cmdPlaces400.Visible = False
    
End Sub

Private Sub cmdPlaces600_Click()
    
    'Shows and hides the forms
    frmPlaces600.Show
    frmGameBoard.Hide
    
    'Hides the command button
    cmdPlaces600.Visible = False
    
End Sub

Private Sub cmdPlaces800_Click()
    
    'Shows and hides the forms
    frmPlaces800.Show
    frmGameBoard.Hide
    
    'Hides the command button
    cmdPlaces800.Visible = False
    
End Sub

Private Sub Form_Activate()
    
    'Setting initial value of wager
    Wager = -100
    
    'Shows and hides the forms
    Select Case Player
        Case Is = 1
            frmKenMoney.Hide
            frmGameBoard.Show
        Case Is = 2
            frmBushMoney.Hide
            frmGameBoard.Show
    End Select
    
    'Shows the introduction only once to the user
    If cmdCartoons1000.Visible = True And cmdCartoons200.Visible = True And cmdCartoons400.Visible = True And cmdCartoons600.Visible = True And cmdCartoons800.Visible = True And cmdHis1000.Visible = True And cmdHis200.Visible = True And cmdHis400.Visible = True And cmdHis600.Visible = True And cmdHis800.Visible = True And cmdLit200.Visible = True And cmdLit400.Visible = True And cmdLit600.Visible = True And cmdLit800.Visible = True And cmdLit1000.Visible = True And cmdMath1000.Visible = True And cmdMath200.Visible = True And cmdMath400.Visible = True And cmdMath600.Visible = True And cmdMath800.Visible = True And cmdof1000.Visible = True And cmdof200.Visible = True And cmdof400.Visible = True And cmdof600.Visible = True And cmdof800.Visible = True And cmdPlaces1000.Visible = True And cmdPlaces200.Visible = True And cmdPlaces400.Visible = True And cmdPlaces600.Visible = True And cmdPlaces800.Visible = True Then
        MsgBox "Remember this is Jeopardy so all answers are actually questions but do not use a question mark!!!" & vbNewLine & vbNewLine & "Now there are five different categories to choose from:" & vbNewLine & "The first is Literature authors, in which you are expected to the name author's full name." & vbNewLine & "The second is History." & vbNewLine & "The third is Derivatives." & vbNewLine & "The fourth is Cartoon Characters, in which you are expected to name the character shown in the form of a question" & vbNewLine & "The fifth is 'P'laces on the Map, in which every question you give starts with a 'P'" & vbNewLine & "The sixth is _of_, in which you fill in the blanks and do not forget 'of' in your question" & vbNewLine & vbNewLine & "Hit ok to begin", , "The Catergories"
    End If
    
    'Brings up the final jeopardy answer when all of the command buttons are gone
    If cmdCartoons1000.Visible = False And cmdCartoons200.Visible = False And cmdCartoons400.Visible = False And cmdCartoons600.Visible = False And cmdCartoons800.Visible = False And cmdHis1000.Visible = False And cmdHis200.Visible = False And cmdHis400.Visible = False And cmdHis600.Visible = False And cmdHis800.Visible = False And cmdLit200.Visible = False And cmdLit400.Visible = False And cmdLit600.Visible = False And cmdLit800.Visible = False And cmdLit1000.Visible = False And cmdMath1000.Visible = False And cmdMath200.Visible = False And cmdMath400.Visible = False And cmdMath600.Visible = False And cmdMath800.Visible = False And cmdof1000.Visible = False And cmdof200.Visible = False And cmdof400.Visible = False And cmdof600.Visible = False And cmdof800.Visible = False And cmdPlaces1000.Visible = False And cmdPlaces200.Visible = False And cmdPlaces400.Visible = False And cmdPlaces600.Visible = False And cmdPlaces800.Visible = False Then
        frmGameBoard.Hide
        MsgBox "Now it is time for the answer of all answers.  You have reached the final jeopardy round." & vbNewLine & "You can wager as much of your winnings as you want but no more." & vbNewLine & "The category is Computer Science Professors.", , "Final Jeopardy"
            'Allows the user to set a wager if he/she has winnings
            If Winnings > 0 Then
                Do Until Wager >= 0 And Wager <= Winnings And Wager / 100 = Int(Wager / 100)
                    Wager = InputBox("Please enter your wager amount." & vbNewLine & "Remember that it must be less than or equal to your winnings, a multiple of 100, and positive!!!", "Wager Amount")
                    If Wager < 0 Or Wager > Winnings Or Wager / 100 <> Int(Wager / 100) Then
                        MsgBox "You have entered an invalid wager.  Please try again!!!", , "Cheater!!!"
                    End If
                Loop
                frmFinalJeopardy.Show
            Else
                MsgBox "You have no winnings, but you can do the final jeopardy round for fun!!!", , "No Winnings"
                frmFinalJeopardy.Show
            End If
    End If

End Sub
