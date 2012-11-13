VERSION 5.00
Begin VB.Form frmDungeon1 
   BackColor       =   &H00404040&
   Caption         =   "A Dank Dungeon"
   ClientHeight    =   3990
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   ScaleHeight     =   3990
   ScaleWidth      =   7275
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDungeon1 
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   3000
      Width           =   2895
   End
   Begin VB.CommandButton cmdSubmit1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Submit"
      Height          =   375
      Left            =   1080
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton cmdLook1 
      BackColor       =   &H00C0C000&
      Caption         =   "Look Around"
      Height          =   375
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton cmdQuit1 
      BackColor       =   &H000000FF&
      Caption         =   "Quit Program"
      Height          =   375
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3480
      Width           =   2175
   End
   Begin VB.PictureBox picDungeon1 
      Height          =   2655
      Left            =   1080
      ScaleHeight     =   2595
      ScaleWidth      =   5235
      TabIndex        =   0
      Top             =   120
      Width           =   5295
   End
End
Attribute VB_Name = "frmDungeon1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim back As Boolean
Dim D As Integer
'this form is the first room in the main game.  it starts the player in a cell
'and the player must figure out how to get out by using some bones to pick the lock.

Private Sub cmdLook1_Click()
'this provides a description of the room based on what conditions have been fulfilled
'by the player's actions.
picDungeon1.Cls
If bone = False And cagelock = True Then
    picDungeon1.Print "Ye find yeself in yon dungeon.  Ye doth not knowest"
    picDungeon1.Print "how ye got there.  Ye art held prisoner in a rusty"
    picDungeon1.Print "olde cage.  Ye hath better find a way to escape from it."
    picDungeon1.Print "You look around the room, noting its consistent style of"
    picDungeon1.Print "dankity and dreariness.  There are some BONES on the floor of"
    picDungeon1.Print "your cage.  The door to the cage is locked, but the door to"
    picDungeon1.Print "the NORTH beyond it is wide open.  You just need to escape."
End If

If bone = True And cagelock = True Then
    picDungeon1.Print "Well, you've got some nasty olde BONES, and the cage door is"
    picDungeon1.Print "still locked. Whatcha gonna do?"
End If

If cagelock = False Then
    picDungeon1.Print "Things seem much nicer on the outside of the olde cage, but"
    picDungeon1.Print "you still need to get out.  You've got dinner in the oven at home."
    picDungeon1.Print "There is an open door to the NORTH."
End If

End Sub

Private Sub cmdQuit1_Click()
'ends the program
End
End Sub

Private Sub cmdSubmit1_Click()
'checks to see if the users submition mathces any of the actions stored in the
'actions array for this form.
D = 0
submit = txtDungeon1.Text
picDungeon1.Cls

    If LCase(submit) = DungeonActions1(1) And bone = False Then
        bone = True
        picDungeon1.Print "You pick up some gross/cool olde bones. But how"
        picDungeon1.Print "will you impress your friends out in this cage, thou must escape!"
    End If
    If LCase(submit) = DungeonActions1(2) And bone = True And cagelock = True Then
        cagelock = False
        picDungeon1.Print "You try using the bones to pick the lock, some break"
        picDungeon1.Print "but eventually one gets the door open. Now your"
        picDungeon1.Print "using your noggin'!"
    End If
    If LCase(submit) = DungeonActions1(3) And cagelock = False Then
        MsgBox "You leave the cell room behind you and proceed to the next room."
        back = True
        frmDungeon1.Hide
        frmDungeon2.Show
    End If

End Sub

Private Sub Form_Load()
'this sets the variables for when the form is first loaded.
If back = False Then
    cagelock = True
    bone = False
End If
End Sub


