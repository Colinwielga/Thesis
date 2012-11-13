VERSION 5.00
Begin VB.Form frmNonSmokingDoubleDiagram 
   Caption         =   "Smoking Hotel Diagram"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   Picture         =   "frmNonSmokingDoubleDiagram.frx":0000
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdShowAvailable 
      Caption         =   "Show Available Rooms"
      Height          =   975
      Left            =   2640
      TabIndex        =   42
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Return to Previous Page"
      Height          =   975
      Left            =   4680
      TabIndex        =   41
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton CmdMasterSuite40 
      BackColor       =   &H0080FF80&
      Caption         =   "40"
      Enabled         =   0   'False
      Height          =   1455
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton CmdMasterSuite39 
      Caption         =   "39"
      Enabled         =   0   'False
      Height          =   1455
      Left            =   7800
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   3000
      UseMaskColor    =   -1  'True
      Width           =   855
   End
   Begin VB.CommandButton CmdSmallSuite38 
      BackColor       =   &H0080FF80&
      Caption         =   "38"
      Enabled         =   0   'False
      Height          =   495
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   6840
      Width           =   855
   End
   Begin VB.CommandButton CmdSmallSuite37 
      BackColor       =   &H0080FF80&
      Caption         =   "37"
      Enabled         =   0   'False
      Height          =   495
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   6360
      Width           =   855
   End
   Begin VB.CommandButton CmdSmallSuite36 
      BackColor       =   &H0080FF80&
      Caption         =   "36"
      Enabled         =   0   'False
      Height          =   495
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton CmdSmallSuite35 
      Caption         =   "35"
      Enabled         =   0   'False
      Height          =   495
      Left            =   7800
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   2520
      UseMaskColor    =   -1  'True
      Width           =   855
   End
   Begin VB.CommandButton CmdSmallSuite34 
      Caption         =   "34"
      Enabled         =   0   'False
      Height          =   495
      Left            =   7800
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   2040
      UseMaskColor    =   -1  'True
      Width           =   855
   End
   Begin VB.CommandButton CmdSmallSuite33 
      Caption         =   "33"
      Enabled         =   0   'False
      Height          =   495
      Left            =   7800
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   1560
      UseMaskColor    =   -1  'True
      Width           =   855
   End
   Begin VB.CommandButton CmdKing32 
      BackColor       =   &H0080FF80&
      Caption         =   "32"
      Enabled         =   0   'False
      Height          =   495
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   7320
      Width           =   495
   End
   Begin VB.CommandButton CmdKing31 
      BackColor       =   &H0080FF80&
      Caption         =   "31"
      Enabled         =   0   'False
      Height          =   495
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   7320
      Width           =   495
   End
   Begin VB.CommandButton CmdKing30 
      BackColor       =   &H0080FF80&
      Caption         =   "30"
      Enabled         =   0   'False
      Height          =   495
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   7320
      Width           =   495
   End
   Begin VB.CommandButton CmdQueen29 
      BackColor       =   &H0080FF80&
      Caption         =   "29"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   7320
      Width           =   495
   End
   Begin VB.CommandButton CmdQueen28 
      BackColor       =   &H0080FF80&
      Caption         =   "28"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   7320
      Width           =   495
   End
   Begin VB.CommandButton CmdQueen27 
      BackColor       =   &H0080FF80&
      Caption         =   "27"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   7320
      Width           =   495
   End
   Begin VB.CommandButton CmdQueen26 
      BackColor       =   &H0080FF80&
      Caption         =   "26"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   7320
      Width           =   495
   End
   Begin VB.CommandButton CmdDouble25 
      BackColor       =   &H0080FF80&
      Caption         =   "25"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4080
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   7320
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton CmdDouble24 
      BackColor       =   &H0080FF80&
      Caption         =   "24"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3720
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   7320
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton CmdDouble23 
      BackColor       =   &H0080FF80&
      Caption         =   "23"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3360
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   7320
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton CmdDouble22 
      BackColor       =   &H0080FF80&
      Caption         =   "22"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3000
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   7320
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton CmdDouble21 
      BackColor       =   &H0080FF80&
      Caption         =   "21"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2640
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   7320
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton CmdDouble20 
      BackColor       =   &H0080FF80&
      Caption         =   "20"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2280
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   7320
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton CmdDouble19 
      BackColor       =   &H0080FF80&
      Caption         =   "19"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1920
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   7320
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton CmdDouble18 
      BackColor       =   &H0080FF80&
      Caption         =   "18"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1560
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7320
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton CmdDouble17 
      BackColor       =   &H0080FF80&
      Caption         =   "17"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1200
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7320
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton CmdKing16 
      Caption         =   "16"
      Enabled         =   0   'False
      Height          =   495
      Left            =   7320
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1080
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CommandButton CmdKing15 
      Caption         =   "15"
      Enabled         =   0   'False
      Height          =   495
      Left            =   6840
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1080
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CommandButton CmdKing14 
      Caption         =   "14"
      Enabled         =   0   'False
      Height          =   495
      Left            =   6360
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1080
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CommandButton cmdQueen13 
      Caption         =   "13"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5880
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1080
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CommandButton CmdQueen12 
      Caption         =   "12"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5400
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1080
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CommandButton cmdQueen11 
      Caption         =   "11"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4920
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1080
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CommandButton CmdQueen10 
      Caption         =   "10"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4440
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1080
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CommandButton cmdDouble9 
      Caption         =   "9"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4080
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1080
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton CmdDouble8 
      Caption         =   "8"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3720
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1080
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton CmdDouble7 
      Caption         =   "7"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3360
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1080
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton CmdDouble6 
      Caption         =   "6"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3000
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1080
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton CmdDouble5 
      Caption         =   "5"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2640
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1080
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdDouble4 
      Caption         =   "4"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2280
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1080
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdDouble3 
      Caption         =   "3"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1920
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1080
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton CmdDouble2 
      Caption         =   "2"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1560
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1080
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdDouble1 
      Caption         =   "1"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1200
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1080
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.Label lblChoose 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "Click ""Show Available Rooms"" when this page loads!"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   960
      TabIndex        =   43
      Top             =   3600
      Width           =   6495
   End
   Begin VB.Label lblNonSmokingRooms 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Non-Smoking Rooms"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   40
      Top             =   4200
      Width           =   3375
   End
End
Attribute VB_Name = "frmNonSmokingDoubleDiagram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project: Hotel Checkin
'Form: NonSmoking Rooms
'Authors: Ellen Jansen & Stuart Van Ess
'Date: March 28, 2008
'Purpose:   This is the page where we choose which Non-smoking Room we would
'           like to place the customer in. We can also show what rooms are
'           available and which are not.
    



Option Explicit

Private Sub cmdBack_Click()
'If no rooms are available, customers have the option to return and look at
'different rooms.
    frmRoomSize.Show
    frmNonSmokingDoubleDiagram.Hide
End Sub
'Each of the buttons with a room number on it do 3 things.
'1.) when clicked they change from green to red, indicating the room is occupied
'2.) when clicked they become no longer enabled, so the user can not click it
'3.) the checkin information menu is shown.

Private Sub CmdDouble17_Click()
    CmdDouble17.BackColor = &HFF&
    CmdDouble17.Enabled = False
    frmCheckin.Show
    
End Sub

Private Sub CmdDouble18_Click()
    CmdDouble18.BackColor = &HFF&
    CmdDouble18.Enabled = False
    frmCheckin.Show
End Sub

Private Sub CmdDouble19_Click()
    CmdDouble19.BackColor = &HFF&
    CmdDouble19.Enabled = False
    frmCheckin.Show
End Sub

Private Sub CmdDouble20_Click()
    CmdDouble20.BackColor = &HFF&
    CmdDouble20.Enabled = False
    frmCheckin.Show
End Sub

Private Sub CmdDouble21_Click()
    CmdDouble21.BackColor = &HFF&
    CmdDouble21.Enabled = False
    frmCheckin.Show
End Sub

Private Sub CmdDouble22_Click()
    CmdDouble22.BackColor = &HFF&
    CmdDouble22.Enabled = False
    frmCheckin.Show
End Sub

Private Sub CmdDouble23_Click()
    CmdDouble23.BackColor = &HFF&
    CmdDouble23.Enabled = False
    frmCheckin.Show
End Sub

Private Sub CmdDouble24_Click()
    CmdDouble24.BackColor = &HFF&
    CmdDouble24.Enabled = False
    frmCheckin.Show
End Sub

Private Sub CmdDouble25_Click()
    CmdDouble25.BackColor = &HFF&
    CmdDouble25.Enabled = False
    frmCheckin.Show
End Sub

Private Sub CmdKing30_Click()
    CmdKing30.BackColor = &HFF&
    CmdKing30.Enabled = False
    frmCheckin.Show
End Sub

Private Sub CmdKing31_Click()
    CmdKing31.BackColor = &HFF&
    CmdKing31.Enabled = False
    frmCheckin.Show
End Sub

Private Sub CmdKing32_Click()
    CmdKing32.BackColor = &HFF&
    CmdKing32.Enabled = False
    frmCheckin.Show
End Sub

Private Sub CmdMasterSuite40_Click()
    CmdMasterSuite40.BackColor = &HFF&
    CmdMasterSuite40.Enabled = False
    frmCheckin.Show
End Sub

Private Sub CmdQueen26_Click()
    CmdQueen26.BackColor = &HFF&
    CmdQueen26.Enabled = False
    frmCheckin.Show
End Sub

Private Sub CmdQueen27_Click()
    CmdQueen27.BackColor = &HFF&
    CmdQueen27.Enabled = False
    frmCheckin.Show
End Sub

Private Sub CmdQueen28_Click()
    CmdQueen28.BackColor = &HFF&
    CmdQueen28.Enabled = False
    frmCheckin.Show
End Sub

Private Sub CmdQueen29_Click()
    CmdQueen29.BackColor = &HFF&
    CmdQueen29.Enabled = False
    frmCheckin.Show
End Sub


'this button makes it so only the "King", or only the "Queen", etc. rooms are
'available when the customer selects them accordingly.
Private Sub cmdShowAvailable_Click()
    If Cozy = "Double" Then
        CmdDouble17.Enabled = True
        CmdDouble18.Enabled = True
        CmdDouble19.Enabled = True
        CmdDouble20.Enabled = True
        CmdDouble21.Enabled = True
        CmdDouble22.Enabled = True
        CmdDouble23.Enabled = True
        CmdDouble24.Enabled = True
        CmdDouble25.Enabled = True
        CmdQueen26.Enabled = False
        CmdQueen27.Enabled = False
        CmdQueen28.Enabled = False
        CmdQueen29.Enabled = False
        CmdKing30.Enabled = False
        CmdKing31.Enabled = False
        CmdKing32.Enabled = False
        CmdSmallSuite36.Enabled = False
        CmdSmallSuite37.Enabled = False
        CmdSmallSuite38.Enabled = False
        CmdMasterSuite40.Enabled = False
    End If
    If Cozy = "Queen" Then
        CmdDouble17.Enabled = False
        CmdDouble18.Enabled = False
        CmdDouble19.Enabled = False
        CmdDouble20.Enabled = False
        CmdDouble21.Enabled = False
        CmdDouble22.Enabled = False
        CmdDouble23.Enabled = False
        CmdDouble24.Enabled = False
        CmdDouble25.Enabled = False
        CmdQueen26.Enabled = True
        CmdQueen27.Enabled = True
        CmdQueen28.Enabled = True
        CmdQueen29.Enabled = True
        CmdKing30.Enabled = False
        CmdKing31.Enabled = False
        CmdKing32.Enabled = False
        CmdSmallSuite36.Enabled = False
        CmdSmallSuite37.Enabled = False
        CmdSmallSuite38.Enabled = False
        CmdMasterSuite40.Enabled = False
    End If
    If Cozy = "King" Then
        CmdDouble17.Enabled = False
        CmdDouble18.Enabled = False
        CmdDouble19.Enabled = False
        CmdDouble20.Enabled = False
        CmdDouble21.Enabled = False
        CmdDouble22.Enabled = False
        CmdDouble23.Enabled = False
        CmdDouble24.Enabled = False
        CmdDouble25.Enabled = False
        CmdQueen26.Enabled = False
        CmdQueen27.Enabled = False
        CmdQueen28.Enabled = False
        CmdQueen29.Enabled = False
        CmdKing30.Enabled = True
        CmdKing31.Enabled = True
        CmdKing32.Enabled = True
        CmdSmallSuite36.Enabled = False
        CmdSmallSuite37.Enabled = False
        CmdSmallSuite38.Enabled = False
        CmdMasterSuite40.Enabled = False
    End If
    
    If Cozy = "SmallSuite" Then
        CmdDouble17.Enabled = False
        CmdDouble18.Enabled = False
        CmdDouble19.Enabled = False
        CmdDouble20.Enabled = False
        CmdDouble21.Enabled = False
        CmdDouble22.Enabled = False
        CmdDouble23.Enabled = False
        CmdDouble24.Enabled = False
        CmdDouble25.Enabled = False
        CmdQueen26.Enabled = False
        CmdQueen27.Enabled = False
        CmdQueen28.Enabled = False
        CmdQueen29.Enabled = False
        CmdKing30.Enabled = False
        CmdKing31.Enabled = False
        CmdKing32.Enabled = False
        CmdSmallSuite36.Enabled = True
        CmdSmallSuite37.Enabled = True
        CmdSmallSuite38.Enabled = True
        CmdMasterSuite40.Enabled = False
    End If
    If Cozy = "MasterSuite" Then
        CmdDouble17.Enabled = False
        CmdDouble18.Enabled = False
        CmdDouble19.Enabled = False
        CmdDouble20.Enabled = False
        CmdDouble21.Enabled = False
        CmdDouble22.Enabled = False
        CmdDouble23.Enabled = False
        CmdDouble24.Enabled = False
        CmdDouble25.Enabled = False
        CmdQueen26.Enabled = False
        CmdQueen27.Enabled = False
        CmdQueen28.Enabled = False
        CmdQueen29.Enabled = False
        CmdKing30.Enabled = False
        CmdKing31.Enabled = False
        CmdKing32.Enabled = False
        CmdSmallSuite36.Enabled = False
        CmdSmallSuite37.Enabled = False
        CmdSmallSuite38.Enabled = False
        CmdMasterSuite40.Enabled = True
    End If
End Sub

Private Sub CmdSmallSuite36_Click()
    CmdSmallSuite36.BackColor = &HFF&
    CmdSmallSuite36.Enabled = False
    frmCheckin.Show
End Sub

Private Sub CmdSmallSuite37_Click()
    CmdSmallSuite37.BackColor = &HFF&
    CmdSmallSuite37.Enabled = False
    frmCheckin.Show
End Sub

Private Sub CmdSmallSuite38_Click()
    CmdSmallSuite38.BackColor = &HFF&
    CmdSmallSuite38.Enabled = False
    frmCheckin.Show
End Sub

