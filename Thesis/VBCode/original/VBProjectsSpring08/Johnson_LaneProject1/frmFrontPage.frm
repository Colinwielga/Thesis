VERSION 5.00
Begin VB.Form frmFrontPage 
   Caption         =   "NikeTown"
   ClientHeight    =   7845
   ClientLeft      =   2115
   ClientTop       =   1905
   ClientWidth     =   11760
   LinkTopic       =   "Form1"
   Picture         =   "frmFrontPage.frx":0000
   ScaleHeight     =   7845
   ScaleWidth      =   11760
   Begin VB.CommandButton CmdQuit 
      Caption         =   "                        Leave Nike Town"
      Height          =   1215
      Left            =   120
      Picture         =   "frmFrontPage.frx":3227D
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6480
      Width           =   2775
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "                      Welcome to NikeTown"
      Height          =   1215
      Left            =   120
      Picture         =   "frmFrontPage.frx":34DC4
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5040
      Width           =   2775
   End
End
Attribute VB_Name = "frmFrontPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Project name: Nike Town
'Form name: frmFrontPage
'Author: Sean Johnson and Nick Lane
'Date Written: Friday March 14th, 2007
'Objective of form: this project is designed to simulate the actions of a typical sports store
'                   including information on various items, buy option, checkout options, as well
'                   as interesting challenges for the which can be completed for extra discounts.
'                   This form acts as a title page does for an essay. it is the main page which
'                   which allows access to the program. it allows the user to enter the program
'                   and exit the program.


Private Declare Function mciSendString Lib "winmm.dll" Alias _
        "mciSendStringA" (ByVal lpstrCommand As String, ByVal _
        lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal _
        hwndCallback As Long) As Long
        
Private Sub cmdEnter_Click()

        'opens the welcome video
        Const FILE_TO_OPEN = " M:\Johnson_Laneproject1\untitled.wmv"
        Dim strCmdStr As String
        Dim lngReturnVal As Long
        strCmdStr = "play " & FILE_TO_OPEN & " fullscreen " 'tells the computer to play the file
        lngReturnVal = mciSendString(strCmdStr, 0&, 0, 0&) ' starts playing the avi file from the begginning of the file
        lngReturnVal = mciSendString("play FILE_TO_OPEN wait", 0&, 0, 0) ' allows the entire AVI file to play without interruption
        
      MsgBox "WELCOME TO NIKETOWN", , "Welcome" 'gives the user a welcome message

        frmFrontPage.Hide               'hides this form and displays the second form
        frmSecondPage.Show
        
            
End Sub

Private Sub CmdQuit_Click()

'message to thank the user for using the program
MsgBox "Thankyou for shopping at NikeTown. Come Back soon!", , "GOODBYE"
End

End Sub

Private Sub cmdWelcome_Click()

MsgBox "WELCOME TO NIKETOWN", , "Welcome" 'gives the user a welcome message

frmFrontPage.Hide               'hides this form and displays the second form
frmSecondPage.Show

End Sub
