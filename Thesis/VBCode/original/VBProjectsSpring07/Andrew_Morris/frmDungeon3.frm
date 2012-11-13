VERSION 5.00
Begin VB.Form frmDungeon3 
   BackColor       =   &H00800080&
   Caption         =   "The West Wing"
   ClientHeight    =   3990
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   ScaleHeight     =   3990
   ScaleWidth      =   7275
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDungeon3 
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   3000
      Width           =   2895
   End
   Begin VB.CommandButton cmdSubmit3 
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
   Begin VB.CommandButton cmdLook3 
      BackColor       =   &H00C0C000&
      Caption         =   "Look Around"
      Height          =   375
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton cmdQuit3 
      BackColor       =   &H000000FF&
      Caption         =   "Quit Program"
      Height          =   375
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3480
      Width           =   2175
   End
   Begin VB.PictureBox picDungeon3 
      Height          =   2655
      Left            =   1080
      ScaleHeight     =   2595
      ScaleWidth      =   5235
      TabIndex        =   0
      Top             =   120
      Width           =   5295
   End
End
Attribute VB_Name = "frmDungeon3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'this room is another object oriented puzzle.  the user must use various objects
'available in the room to press a button on the ceiling which opens the way to the
'north door in the main chamber.

Private Sub cmdLook3_Click()
'this button provides a description of the current status of the room based on what
'appropriate conditions have been met.
picDungeon3.Cls
If yarn = False Then
    picDungeon3.Print "This seems to be a practice room of some kind.  There is some"
    picDungeon3.Print "YARN on a nearby shelf. The room consists of a series of platforms,"
    picDungeon3.Print "none of which you can reach. Some DARTS lie on the next platform"
    picDungeon3.Print "and a pressure button is on the ceiling. Better get to work."
End If
If yarn = True And darts = False Then
    picDungeon3.Print "You've got the YARN, but the button still eludes you. If only"
    picDungeon3.Print "you could throw something at it, like those DARTS...."
End If
If darts = True And puzzle1 = False Then
    picDungeon3.Print "You stare at the button, wondering how in the world you could activate it."
End If
If puzzle1 = True Then
    picDungeon3.Print "The room now looks bare, with the exception of the floating platforms and"
    picDungeon3.Print "the odd-looking button on the ceiling.  Nothing left to do here."
End If
picDungeon3.Print "The door leading back to the central chamber is to the EAST."
End Sub

Private Sub cmdQuit3_Click()
'the ends the program
End
End Sub

Private Sub cmdSubmit3_Click()
'this checks to see if the user's submission matches the appropriate actions for
'this room as stored in this rooms actions array.
picDungeon3.Cls
submit = txtDungeon3.Text
If LCase(submit) = DungeonActions3(2) And yarn = False Then
    yarn = True
    picDungeon3.Print "You grab the YARN, hoping to make use of it in hitting the ceiling button."
End If
If LCase(submit) = DungeonActions3(3) And darts = False Then
    darts = True
    picDungeon3.Print "You make the yarn into a a lasso and use it to pull the DARTS down."
    picDungeon3.Print "You then proceed to pick up the DARTS."
End If
If LCase(submit) = DungeonActions3(4) And puzzle1 = False Then
    puzzle1 = True
    picDungeon3.Print "You hurl a dart at the button. Bullseye!"
    picDungeon3.Print "You can hear something move in the central chamber."
End If
If LCase(submit) = DungeonActions3(4) And puzzle1 = True Then
    picDungeon3.Print "You decide to keep the darts; they're pretty cool"
End If
If LCase(submit) = DungeonActions3(1) And puzzle1 = False Then
    MsgBox "You head back into the main chamber, with the button eluding you.", , "Hmmmm"
    frmDungeon3.Hide
    frmDungeon2.Show
End If
If LCase(submit) = DungeonActions3(1) And puzzle1 = True Then
    MsgBox "You head back to the main chamber to see what has changed.", , "Great Success"
    frmDungeon3.Hide
    frmDungeon2.Show
End If
End Sub

Private Sub Form_Load()
'this sets the conditions for the room upon entering it for the first time
If puzzle1 = False Then
    yarn = False
    darts = False
End If
End Sub
