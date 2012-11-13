VERSION 5.00
Begin VB.Form frmFishing 
   BackColor       =   &H000000FF&
   Caption         =   "Go to Fishing Page"
   ClientHeight    =   6990
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10485
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form1"
   Picture         =   "frmFishing.frx":0000
   ScaleHeight     =   6990
   ScaleWidth      =   10485
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdYellow 
      BackColor       =   &H0000FFFF&
      Height          =   975
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton cmdGreen 
      BackColor       =   &H00008000&
      Height          =   975
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton cmdBlack 
      BackColor       =   &H00000000&
      Height          =   975
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton cmdRed 
      BackColor       =   &H000000FF&
      Height          =   975
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox txtFish 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   720
      TabIndex        =   4
      Text            =   "Pick a color to see what fish you are"
      Top             =   1320
      Width           =   4695
   End
   Begin VB.CommandButton cmdBlue 
      BackColor       =   &H00FF0000&
      Height          =   975
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton cmdHowOften 
      BackColor       =   &H00C0C0C0&
      Caption         =   "How many days a year do you fish?"
      Height          =   855
      Left            =   720
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4080
      Width           =   2655
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Return to Main Page"
      Height          =   855
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5880
      Width           =   2655
   End
   Begin VB.TextBox txtTitle 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   720
      TabIndex        =   0
      Text            =   "Fishing in Minnesota"
      Top             =   240
      Width           =   4695
   End
End
Attribute VB_Name = "frmFishing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Minnesota DNR
'Fishing
'Andrew Forlit and Casey Orthaus
'October 19th, 2009
'this is the fishing page, it tells you what fish you are by what color you select and
'this page tells you if you have a fishing problem

Private Sub cmdBlack_Click()

MsgBox "You are a Catfish", , "Black = Catfish"

frmFishing.Hide
frmCatFish.Show

End Sub

Private Sub cmdBlue_Click()

MsgBox "You are a Walleye", , "Blue = Walleye"

frmWalleye.Show
frmFishing.Hide

End Sub

Private Sub cmdGreen_Click()

MsgBox "You are a Crappie", , "Green = Crappie"

frmFishing.Hide
frmCrappie.Show

End Sub

Private Sub cmdHowOften_Click()

'declaring variables
Dim days As Integer

'input box
days = InputBox("How many days a year do you Fish?")

Select Case days
    Case 0 To 20
        MsgBox "You need to get out and fish more"
    Case 21 To 80
        MsgBox "You are beginning to understand how great fishing is"
    Case 81 To 150
        MsgBox "You know how relaxing fishing is and love it"
    Case 151 To 200
        MsgBox "You might have a fishing problem"
    Case 200 To 365
        MsgBox "You have a fishing problem"
    Case Else
        MsgBox "Entered an invalid number of days", , "Error"
End Select
    

End Sub

Private Sub cmdRed_Click()

MsgBox "You are a Northern Pike", , "Red = Northern Pike"

frmFishing.Hide
frmNorthern.Show

End Sub

Private Sub cmdReturn_Click()

frmDNR.Show
frmFishing.Hide

End Sub


Private Sub cmdYellow_Click()

MsgBox "You are Largemouth Bass", , "Yellow = Largemouth Bass"

frmFishing.Hide
frmBass.Show

End Sub

