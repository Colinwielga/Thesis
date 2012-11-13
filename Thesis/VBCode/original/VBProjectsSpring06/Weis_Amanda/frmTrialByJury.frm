VERSION 5.00
Begin VB.Form frmTrialByJury 
   BackColor       =   &H00000000&
   Caption         =   "Trial By Jury"
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9045
   LinkTopic       =   "Form1"
   Picture         =   "frmTrialByJury.frx":0000
   ScaleHeight     =   7500
   ScaleWidth      =   9045
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCastList 
      Caption         =   "Cast List"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6240
      TabIndex        =   2
      ToolTipText     =   "To learn more about the characters and see the cast list of the show, click here."
      Top             =   360
      Width           =   2415
   End
   Begin VB.CommandButton cmdSynopsis 
      Caption         =   "Synopsis"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   1
      ToolTipText     =   "Click to learn the plot of Trial By Jury."
      Top             =   4200
      Width           =   2175
   End
   Begin VB.CommandButton cmdGoBack 
      Caption         =   "Go Back"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      TabIndex        =   0
      ToolTipText     =   "Click to go back to previous form."
      Top             =   6240
      Width           =   2175
   End
   Begin VB.Label lblTrialByJury 
      BackColor       =   &H80000007&
      Caption         =   "Trial By Jury"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1215
      Left            =   480
      TabIndex        =   4
      Top             =   6120
      Width           =   4095
   End
   Begin VB.Label lblDesign 
      BackColor       =   &H00000000&
      Caption         =   "Designed By Amanda Weis"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   7440
      TabIndex        =   3
      Top             =   6720
      Width           =   1335
   End
End
Attribute VB_Name = "frmTrialByJury"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    'create button to proceed to Cast List form
Private Sub cmdCastList_Click()
    frmTrialByJuryCastList.Show
    frmTrialByJury.Hide
End Sub
    'create button to go back to previous form
Private Sub cmdGoBack_Click()
    frmTrialByJury.Hide
    frmOPERA.Show
End Sub
    'create button to display a message box with a synopsis
Private Sub cmdSynopsis_Click()
        MsgBox "Trial by Jury  is a spoof of the British legal system.  The beautiful defendant, Angelina, is suing her would-be husband, Edwin for a breach-of-promise.   The court consists of terribly biased participants, including an all male jury and a less-than virtuous judge, who are immediately smitten with Angelina and insensitive to the pleas of Edwin.  ", , "Trial By Jury Synopsis"
End Sub


