VERSION 5.00
Begin VB.Form frmSmokingDoubleDiagram 
   Caption         =   "Smoking Hotel Diagram"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   Picture         =   "frmSmokingDoubleDiagram.frx":0000
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdShowAvailable 
      Caption         =   "Show Available Rooms"
      Height          =   975
      Left            =   2520
      TabIndex        =   42
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Return to Previous Page"
      Height          =   975
      Left            =   4320
      TabIndex        =   41
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton CmdMasterSuite40 
      Caption         =   "40"
      Enabled         =   0   'False
      Height          =   1455
      Left            =   7800
      TabIndex        =   39
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton CmdMasterSuite39 
      BackColor       =   &H0080FF80&
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
      Caption         =   "38"
      Enabled         =   0   'False
      Height          =   495
      Left            =   7800
      TabIndex        =   37
      Top             =   6840
      Width           =   855
   End
   Begin VB.CommandButton CmdSmallSuite37 
      Caption         =   "37"
      Enabled         =   0   'False
      Height          =   495
      Left            =   7800
      TabIndex        =   36
      Top             =   6360
      Width           =   855
   End
   Begin VB.CommandButton CmdSmallSuite36 
      Caption         =   "36"
      Enabled         =   0   'False
      Height          =   495
      Left            =   7800
      TabIndex        =   35
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton CmdSmallSuite35 
      BackColor       =   &H0080FF80&
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
      BackColor       =   &H0080FF80&
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
      BackColor       =   &H0080FF80&
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
      Caption         =   "32"
      Enabled         =   0   'False
      Height          =   495
      Left            =   7320
      TabIndex        =   31
      Top             =   7320
      Width           =   495
   End
   Begin VB.CommandButton CmdKing31 
      Caption         =   "31"
      Enabled         =   0   'False
      Height          =   495
      Left            =   6840
      TabIndex        =   30
      Top             =   7320
      Width           =   495
   End
   Begin VB.CommandButton CmdKing30 
      Caption         =   "30"
      Enabled         =   0   'False
      Height          =   495
      Left            =   6360
      TabIndex        =   29
      Top             =   7320
      Width           =   495
   End
   Begin VB.CommandButton CmdKing29 
      Caption         =   "29"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5880
      TabIndex        =   28
      Top             =   7320
      Width           =   495
   End
   Begin VB.CommandButton CmdKing28 
      Caption         =   "28"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5400
      TabIndex        =   27
      Top             =   7320
      Width           =   495
   End
   Begin VB.CommandButton CmdKing27 
      Caption         =   "27"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4920
      TabIndex        =   26
      Top             =   7320
      Width           =   495
   End
   Begin VB.CommandButton CmdKing26 
      Caption         =   "26"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4440
      TabIndex        =   25
      Top             =   7320
      Width           =   495
   End
   Begin VB.CommandButton CmdDouble25 
      Caption         =   "25"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4080
      TabIndex        =   24
      Top             =   7320
      Width           =   375
   End
   Begin VB.CommandButton CmdDouble24 
      Caption         =   "24"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3720
      TabIndex        =   23
      Top             =   7320
      Width           =   375
   End
   Begin VB.CommandButton CmdDouble23 
      Caption         =   "23"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3360
      TabIndex        =   22
      Top             =   7320
      Width           =   375
   End
   Begin VB.CommandButton CmdDouble22 
      Caption         =   "22"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3000
      TabIndex        =   21
      Top             =   7320
      Width           =   375
   End
   Begin VB.CommandButton CmdDouble21 
      Caption         =   "21"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2640
      TabIndex        =   20
      Top             =   7320
      Width           =   375
   End
   Begin VB.CommandButton CmdDouble20 
      Caption         =   "20"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2280
      TabIndex        =   19
      Top             =   7320
      Width           =   375
   End
   Begin VB.CommandButton CmdDouble19 
      Caption         =   "19"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1920
      TabIndex        =   18
      Top             =   7320
      Width           =   375
   End
   Begin VB.CommandButton CmdDouble18 
      Caption         =   "18"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1560
      TabIndex        =   17
      Top             =   7320
      Width           =   375
   End
   Begin VB.CommandButton CmdDouble17 
      Caption         =   "17"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1200
      TabIndex        =   16
      Top             =   7320
      Width           =   375
   End
   Begin VB.CommandButton CmdKing16 
      BackColor       =   &H0080FF80&
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
      BackColor       =   &H0080FF80&
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
      BackColor       =   &H0080FF80&
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
      BackColor       =   &H0080FF80&
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
      BackColor       =   &H0080FF80&
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
      BackColor       =   &H0080FF80&
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
      BackColor       =   &H0080FF80&
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
      BackColor       =   &H0080FF80&
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
      BackColor       =   &H0080FF80&
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
      BackColor       =   &H0080FF80&
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
      BackColor       =   &H0080FF80&
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
      BackColor       =   &H0080FF80&
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
      BackColor       =   &H0080FF80&
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
      BackColor       =   &H0080FF80&
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
      BackColor       =   &H0080FF80&
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
      BackColor       =   &H0080FF80&
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
   Begin VB.Label lblDirections 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Click ""Show Available Rooms"" When This Form Loads"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   840
      TabIndex        =   43
      Top             =   3960
      Width           =   6855
   End
   Begin VB.Label lblSmokingRooms 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Smoking Rooms"
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
      Top             =   4440
      Width           =   2535
   End
End
Attribute VB_Name = "frmSmokingDoubleDiagram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project: Hotel Checkin
'Form: Smoking Rooms
'Authors: Ellen Jansen & Stuart Van Ess
'Date: March 28, 2008
'Purpose:   This is the page where we choose which smoking Room we would
'           like to place the customer in. We can also show what rooms are
'           available and which are not.
    
'*******************************************************************************
'THE NOTES FOR THIS ARE EXACTLY IDENTICAL AS THE NOTES FOR THE NON-SMOKING ROOMS
'*******************************************************************************


Option Explicit
Private Sub cmdBack_Click()
    frmRoomSize.Show
    frmSmokingDoubleDiagram.Hide
End Sub

Private Sub cmdDouble1_Click()
    cmdDouble1.BackColor = &HFF&
    cmdDouble1.Enabled = False
    frmCheckin.Show
End Sub

Private Sub CmdDouble2_Click()
    CmdDouble2.BackColor = &HFF&
    CmdDouble2.Enabled = False
    frmCheckin.Show
End Sub

Private Sub cmdDouble3_Click()
    cmdDouble3.BackColor = &HFF&
    cmdDouble3.Enabled = False
    frmCheckin.Show
End Sub

Private Sub cmdDouble4_Click()
    cmdDouble4.BackColor = &HFF&
    cmdDouble4.Enabled = False
    frmCheckin.Show
End Sub

Private Sub CmdDouble5_Click()
    CmdDouble5.BackColor = &HFF&
    CmdDouble5.Enabled = False
    frmCheckin.Show
End Sub

Private Sub CmdDouble6_Click()
    CmdDouble6.BackColor = &HFF&
    CmdDouble6.Enabled = False
    frmCheckin.Show
End Sub

Private Sub CmdDouble7_Click()
    CmdDouble7.BackColor = &HFF&
    CmdDouble7.Enabled = False
    frmCheckin.Show
End Sub

Private Sub CmdDouble8_Click()
    CmdDouble8.BackColor = &HFF&
    CmdDouble8.Enabled = False
    frmCheckin.Show
End Sub

Private Sub cmdDouble9_Click()
    cmdDouble9.BackColor = &HFF&
    cmdDouble9.Enabled = False
    frmCheckin.Show
End Sub

Private Sub CmdMasterSuite39_Click()
    CmdMasterSuite39.BackColor = &HFF&
    CmdMasterSuite39.Enabled = False
    frmCheckin.Show
End Sub

Private Sub CmdQueen10_Click()
    CmdQueen10.BackColor = &HFF&
    CmdQueen10.Enabled = False
    frmCheckin.Show
End Sub

Private Sub cmdQueen11_Click()
    cmdQueen11.BackColor = &HFF&
    cmdQueen11.Enabled = False
    frmCheckin.Show
End Sub

Private Sub CmdQueen12_Click()
    CmdQueen12.BackColor = &HFF&
    CmdQueen12.Enabled = False
    frmCheckin.Show
End Sub

Private Sub cmdQueen13_Click()
    cmdQueen13.BackColor = &HFF&
    cmdQueen13.Enabled = False
    frmCheckin.Show
End Sub

Private Sub CmdKing14_Click()
    CmdKing14.BackColor = &HFF&
    CmdKing14.Enabled = False
    frmCheckin.Show
End Sub

Private Sub CmdKing15_Click()
    CmdKing15.BackColor = &HFF&
    CmdKing15.Enabled = False
    frmCheckin.Show
End Sub

Private Sub CmdKing16_Click()
    CmdKing16.BackColor = &HFF&
    CmdKing16.Enabled = False
    frmCheckin.Show
End Sub

Private Sub cmdShowAvailable_Click()
If Cozy = "Double" Then
        cmdDouble1.Enabled = True
        CmdDouble2.Enabled = True
        cmdDouble3.Enabled = True
        cmdDouble4.Enabled = True
        CmdDouble5.Enabled = True
        CmdDouble6.Enabled = True
        CmdDouble7.Enabled = True
        CmdDouble8.Enabled = True
        cmdDouble9.Enabled = True
        CmdQueen10.Enabled = False
        cmdQueen11.Enabled = False
        CmdQueen12.Enabled = False
        cmdQueen13.Enabled = False
        CmdKing14.Enabled = False
        CmdKing15.Enabled = False
        CmdKing16.Enabled = False
        CmdSmallSuite33.Enabled = False
        CmdSmallSuite34.Enabled = False
        CmdSmallSuite35.Enabled = False
        CmdMasterSuite39.Enabled = False
    End If
    If Cozy = "Queen" Then
        cmdDouble1.Enabled = False
        CmdDouble2.Enabled = False
        cmdDouble3.Enabled = False
        cmdDouble4.Enabled = False
        CmdDouble5.Enabled = False
        CmdDouble6.Enabled = False
        CmdDouble7.Enabled = False
        CmdDouble8.Enabled = False
        cmdDouble9.Enabled = False
        CmdQueen10.Enabled = True
        cmdQueen11.Enabled = True
        CmdQueen12.Enabled = True
        cmdQueen13.Enabled = True
        CmdKing14.Enabled = False
        CmdKing15.Enabled = False
        CmdKing16.Enabled = False
        CmdSmallSuite33.Enabled = False
        CmdSmallSuite34.Enabled = False
        CmdSmallSuite35.Enabled = False
        CmdMasterSuite39.Enabled = False
    End If
    If Cozy = "King" Then
        cmdDouble1.Enabled = False
        CmdDouble2.Enabled = False
        cmdDouble3.Enabled = False
        cmdDouble4.Enabled = False
        CmdDouble5.Enabled = False
        CmdDouble6.Enabled = False
        CmdDouble7.Enabled = False
        CmdDouble8.Enabled = False
        cmdDouble9.Enabled = False
        CmdQueen10.Enabled = False
        cmdQueen11.Enabled = False
        CmdQueen12.Enabled = False
        cmdQueen13.Enabled = False
        CmdKing14.Enabled = True
        CmdKing15.Enabled = True
        CmdKing16.Enabled = True
        CmdSmallSuite33.Enabled = False
        CmdSmallSuite34.Enabled = False
        CmdSmallSuite35.Enabled = False
        CmdMasterSuite39.Enabled = False
    End If
    
    If Cozy = "SmallSuite" Then
        cmdDouble1.Enabled = False
        CmdDouble2.Enabled = False
        cmdDouble3.Enabled = False
        cmdDouble4.Enabled = False
        CmdDouble5.Enabled = False
        CmdDouble6.Enabled = False
        CmdDouble7.Enabled = False
        CmdDouble8.Enabled = False
        cmdDouble9.Enabled = False
        CmdQueen10.Enabled = False
        cmdQueen11.Enabled = False
        CmdQueen12.Enabled = False
        cmdQueen13.Enabled = False
        CmdKing14.Enabled = False
        CmdKing15.Enabled = False
        CmdKing16.Enabled = False
        CmdSmallSuite33.Enabled = True
        CmdSmallSuite34.Enabled = True
        CmdSmallSuite35.Enabled = True
        CmdMasterSuite39.Enabled = False
    End If
    If Cozy = "MasterSuite" Then
        cmdDouble1.Enabled = False
        CmdDouble2.Enabled = False
        cmdDouble3.Enabled = False
        cmdDouble4.Enabled = False
        CmdDouble5.Enabled = False
        CmdDouble6.Enabled = False
        CmdDouble7.Enabled = False
        CmdDouble8.Enabled = False
        cmdDouble9.Enabled = False
        CmdQueen10.Enabled = False
        cmdQueen11.Enabled = False
        CmdQueen12.Enabled = False
        cmdQueen13.Enabled = False
        CmdKing14.Enabled = False
        CmdKing15.Enabled = False
        CmdKing16.Enabled = False
        CmdSmallSuite33.Enabled = False
        CmdSmallSuite34.Enabled = False
        CmdSmallSuite35.Enabled = False
        CmdMasterSuite39.Enabled = True
    End If
End Sub

Private Sub CmdSmallSuite33_Click()
    CmdSmallSuite33.BackColor = &HFF&
    CmdSmallSuite33.Enabled = False
    frmCheckin.Show
End Sub

Private Sub CmdSmallSuite34_Click()
    CmdSmallSuite34.BackColor = &HFF&
    CmdSmallSuite34.Enabled = False
    frmCheckin.Show
End Sub

Private Sub CmdSmallSuite35_Click()
    CmdSmallSuite35.BackColor = &HFF&
    CmdSmallSuite35.Enabled = False
    frmCheckin.Show
End Sub

