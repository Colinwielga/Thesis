VERSION 5.00
Begin VB.Form frmaverage 
   BackColor       =   &H00008000&
   BorderStyle     =   0  'None
   Caption         =   "Overall Average"
   ClientHeight    =   8355
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11250
   FillColor       =   &H00FFFFFF&
   FillStyle       =   0  'Solid
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmaverage.frx":0000
   ScaleHeight     =   8355
   ScaleWidth      =   11250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Back To Main Menu"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3038
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4800
      Width           =   5175
   End
   Begin VB.CommandButton cmdaverage 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Click For Average"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   3038
      Picture         =   "frmaverage.frx":38ADE
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1320
      Width           =   5175
   End
End
Attribute VB_Name = "frmaverage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdaverage_Click()

    'declare variables
    Dim TimesArray(1 To 80) As Integer
    Dim Sum As Single
    Dim Average As Single
    Dim CTR As Integer
    Dim Pos As Integer
    'read file into an array
    Open App.Path & "\Times.txt" For Input As #1
    CTR = 0
    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, TimesArray(CTR)
    Loop
    Close #1
    
    'sum the array
    For Pos = 1 To CTR
        Sum = Sum + TimesArray(Pos)
    Next Pos
    
    'compute average and display in a messagebox
    Average = Sum / CTR
    
    MsgBox "The Overall Average 100 Meter Dash Time Is " & FormatNumber(Average, 2), , "Average Time"
    
End Sub

    'create cmd to send the user back to the main menu
Private Sub Command1_Click()
    frmaverage.Hide
    frmwhichfact.Show
End Sub
