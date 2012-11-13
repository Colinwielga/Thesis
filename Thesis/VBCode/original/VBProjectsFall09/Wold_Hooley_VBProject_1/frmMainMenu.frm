VERSION 5.00
Begin VB.Form frmMainMenu 
   BackColor       =   &H00000000&
   Caption         =   "Main Menu"
   ClientHeight    =   8775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13275
   FillColor       =   &H00FFFFFF&
   FillStyle       =   6  'Cross
   ForeColor       =   &H8000000A&
   LinkTopic       =   "Form1"
   Picture         =   "frmMainMenu.frx":0000
   ScaleHeight     =   8775
   ScaleWidth      =   13275
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H80000009&
      Height          =   1455
      Left            =   10560
      Picture         =   "frmMainMenu.frx":38D1E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7080
      Width           =   2415
   End
   Begin VB.CommandButton CmdSortingform 
      BackColor       =   &H0000FFFF&
      Caption         =   "See who the Leaders are Across the Division"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   4215
   End
   Begin VB.CommandButton cmdTeamStats 
      BackColor       =   &H0000FFFF&
      Caption         =   "Click here to view individual North West Ddvion team stats"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4560
      Width           =   2895
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Team As String
'The Northwest Division NHL Stats and Appearal
'Form: frmMainMenu
'Created by Ryan Hooley and Sean Wold
'The Purpose of this project is to demonstrate our ability to use Visual Basic and incorporate themes and fetures we learned in class into our work.

'This subrutine will make the Sort form visible and hide the Main Menu
Private Sub CmdSortingform_Click()
    frmSort.Show
    frmMainMenu.Hide
End Sub

Private Sub cmdTeamStats_Click()
'The purpose of this if then statment is to direct the user to the appropriate form according to their perffered team.
'In the process of navigating through the project the desired form will always replace the form previous.

    Team = InputBox("Select a Team")
        If Team = "Wild" Then
        frmWild.Show
        frmMainMenu.Hide
        
        ElseIf Team = "Avalanche" Then
        frmAvalanche.Show
        frmMainMenu.Hide
        
        ElseIf Team = "Canucks" Then
        frmCanucks.Show
        frmMainMenu.Hide
        
        ElseIf Team = "Oilers" Then
        frmOilers.Show
        frmMainMenu.Hide
        
        ElseIf Team = "Flames" Then
        frmFlames.Show
        frmMainMenu.Hide
        
        Else: MsgBox ("Please choose one of the following: Avalanche, Canucks, Flames, Oilers, or Wild")


        End If
        
End Sub

Private Sub cmdQuit_Click()
End

End Sub
