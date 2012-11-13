VERSION 5.00
Begin VB.Form frmMenu 
   BackColor       =   &H00FFC0FF&
   Caption         =   "Menu"
   ClientHeight    =   8295
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10335
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form3"
   ScaleHeight     =   8295
   ScaleWidth      =   10335
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit"
      Height          =   1095
      Left            =   4020
      TabIndex        =   2
      Top             =   5880
      Width           =   2295
   End
   Begin VB.CommandButton cmdplaygame 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Start a Game of Memory!"
      Default         =   -1  'True
      Height          =   1095
      Left            =   2820
      MaskColor       =   &H80000010&
      TabIndex        =   0
      Top             =   3960
      Width           =   4695
   End
   Begin VB.Label lblwelcome 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      Caption         =   "Welcome to the Game of Memory!"
      ForeColor       =   &H00400040&
      Height          =   615
      Left            =   2040
      TabIndex        =   3
      Top             =   1560
      Width           =   6135
   End
   Begin VB.Label lblmenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      Caption         =   "Main Menu"
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   3840
      TabIndex        =   1
      Top             =   360
      Width           =   2655
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : MemoryGame (Lynn Johnson - VB Project)
'Form Name : frmBestDirector (OscarOddsDirector.frm)
'Author: Gavin Marrin
'Date Written: March 11, 2003
'Purpose of Form: To get Oscar Picks from user and print their
                 'selection with the odds that their selection
                 'will win. Then multiply the odds of all of
                 'there selections to come up with the odds
                 'that all six of their picks will win the
                 'Oscars.

'Option Explicit is a command to force
'the user to explicitly declare all variables
'before they can be used.

Option Explicit

Private Sub cmdplaygame_Click()

    'Move from Menu form to game form
    Dim game As Integer
        game = InputBox("Enter a game number between the numbers 1 and 5")
    
    Select Case game
        Case 1
            frmGame1.Show
            frmMenu.Hide
        Case 2
            frmGame2.Show
            frmMenu.Hide
        Case 3
            frmGame3.Show
            frmMenu.Hide
        Case 4
            FrmGame4.Show
            frmMenu.Hide
        Case 5
            FrmGame5.Show
            frmMenu.Hide
        Case Else
            MsgBox "That number is not between 1 and 5.  Pick another number", , "Error"
    End Select
    
    

    
End Sub

Private Sub cmdquit_Click()
    End
    
End Sub

Private Sub Command1_Click()
    Dim i As Integer
    dim
    pbxresults.Cls
    
    
    For i = 1 To 20
        
    Next i
    
    N = strname(i)
    
    For i = 1 To 20
        pbxresults.Print i; strname(i)
    Next i
    
End Sub

Private Sub Command2_Click()
    For i = 1 To 3
        pbxresults.Print strname(i)
    Next i
End Sub
