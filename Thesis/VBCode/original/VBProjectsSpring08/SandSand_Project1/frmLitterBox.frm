VERSION 5.00
Begin VB.Form frmLitterBox 
   BackColor       =   &H0080C0FF&
   Caption         =   "Litter Boxes"
   ClientHeight    =   10995
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   10995
   ScaleWidth      =   15240
   Begin VB.CommandButton cmdDisposableBox 
      Caption         =   "Disposable"
      BeginProperty Font 
         Name            =   "Hobo Std"
         Size            =   9.75
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12240
      TabIndex        =   10
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CommandButton cmdAutomaticBox 
      Caption         =   "Automatic"
      BeginProperty Font 
         Name            =   "Hobo Std"
         Size            =   9.75
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12240
      TabIndex        =   9
      Top             =   3240
      Width           =   1815
   End
   Begin VB.CommandButton cmdHiddenBox 
      Caption         =   "Hidden"
      BeginProperty Font 
         Name            =   "Hobo Std"
         Size            =   9.75
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12240
      TabIndex        =   8
      Top             =   3960
      Width           =   1815
   End
   Begin VB.CommandButton cmdNoThanks 
      Caption         =   "No Thank You!"
      BeginProperty Font 
         Name            =   "Minion Pro SmBd"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8640
      TabIndex        =   5
      Top             =   9120
      Width           =   1695
   End
   Begin VB.CommandButton cmdMainMenu 
      Caption         =   "Back to Main Menu"
      BeginProperty Font 
         Name            =   "Minion Pro SmBd"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   10920
      TabIndex        =   4
      Top             =   9120
      Width           =   3975
   End
   Begin VB.PictureBox PicResults 
      Height          =   5175
      Left            =   5400
      ScaleHeight     =   5115
      ScaleWidth      =   5355
      TabIndex        =   3
      Top             =   1800
      Width           =   5415
   End
   Begin VB.CommandButton cmdHidden 
      BackColor       =   &H00FF8080&
      Caption         =   " Hidden Litter Box ($115.00)"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5400
      Width           =   2895
   End
   Begin VB.CommandButton cmdAutomatic 
      BackColor       =   &H00FF8080&
      Caption         =   "  Automatic   Scope-Free Litter Box ($50.00)"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3720
      Width           =   2895
   End
   Begin VB.CommandButton cmdDisposal 
      BackColor       =   &H00FF8080&
      Caption         =   "       Disposable               Litter Box         (4 for $10)"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2040
      Width           =   2895
   End
   Begin VB.Label LabInstructions3 
      BackColor       =   &H0080C0FF&
      Caption         =   "Click the type of litter box you wish to purchase:"
      BeginProperty Font 
         Name            =   "Hobo Std"
         Size            =   12
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   855
      Index           =   1
      Left            =   11400
      TabIndex        =   11
      Top             =   1560
      Width           =   3615
   End
   Begin VB.Label labInstructions 
      BackColor       =   &H0080C0FF&
      Caption         =   "To view the types of litter boxes click below:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   720
      TabIndex        =   7
      Top             =   480
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Litter Boxes"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   1215
      Left            =   720
      TabIndex        =   6
      Top             =   8520
      Width           =   6855
   End
End
Attribute VB_Name = "frmLitterBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Sand's Pet Store
'Name of Form: frmLitterBox
'Author: Scott Sand and Kate Sand
'Date Written: March 10, 2008
'Objective:This is where people can view and select litter boxes for their cats.
'Other Comments:

Option Explicit

Private Sub cmdAutomatic_Click()
Open App.Path & "\PicLitterBox.txt" For Input As #1
CTR = 0
Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, Litter(CTR)
Loop
PicResults.Picture = LoadPicture(Litter(2))
Close #1
End Sub

Private Sub cmdAutomaticBox_Click()
MsgBox ("You have purchased an automatic scope-free Litter Box for $50.00.")
HabitatCost = HabitatCost + 50
frmLitterBox.Hide
frmCatToys.Show
End Sub

Private Sub cmdHidden_Click()
Open App.Path & "\PicLitterBox.txt" For Input As #1
CTR = 0
Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, Litter(CTR)
Loop
PicResults.Picture = LoadPicture(Litter(3))
Close #1
End Sub

Private Sub cmdDisposal_Click()
Open App.Path & "\PicLitterBox.txt" For Input As #1
CTR = 0
Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, Litter(CTR)
Loop
PicResults.Picture = LoadPicture(Litter(1))
Close #1
End Sub

Private Sub cmdHiddenBox_Click()
MsgBox ("You have purchased a hidden litter box for $115.00.")
HabitatCost = HabitatCost + 115
frmLitterBox.Hide
frmCatToys.Show
End Sub

Private Sub cmdMainMenu_Click()
frmMainMenu.Show
frmLitterBox.Hide
End Sub

Private Sub cmdNoThanks_Click()
frmLitterBox.Hide
frmCatToys.Show
End Sub

Private Sub cmdDisposableBox_Click()
MsgBox ("You have purchased four disposable litter boxes for $10.00.")
HabitatCost = HabitatCost + 10
frmCatToys.Show
frmLitterBox.Hide
End Sub

