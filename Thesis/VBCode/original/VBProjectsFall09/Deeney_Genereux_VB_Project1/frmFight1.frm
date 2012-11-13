VERSION 5.00
Begin VB.Form frmFight1 
   Caption         =   "Fight The Dragon!"
   ClientHeight    =   5370
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8235
   LinkTopic       =   "Form1"
   ScaleHeight     =   5370
   ScaleWidth      =   8235
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   195
      Left            =   3480
      TabIndex        =   7
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton cmdKungFu 
      Caption         =   "Use Kung Fu Master"
      Height          =   255
      Left            =   5760
      TabIndex        =   6
      Top             =   4560
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdsandwich 
      Caption         =   "Use Sandwich"
      Height          =   255
      Left            =   1920
      TabIndex        =   5
      Top             =   4560
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdSword 
      Caption         =   "Use Sword"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   4560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdToad 
      Caption         =   "Use Toad"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   2880
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdInventory 
      Caption         =   "Look in Inventory"
      Height          =   615
      Left            =   1320
      TabIndex        =   1
      Top             =   1200
      Width           =   2655
   End
   Begin VB.Image img4 
      Height          =   4980
      Left            =   5520
      Picture         =   "frmFight1.frx":0000
      Top             =   0
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Image img3 
      Height          =   1560
      Left            =   240
      Picture         =   "frmFight1.frx":36F0
      Top             =   3360
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Image img2 
      Height          =   2880
      Left            =   1680
      Picture         =   "frmFight1.frx":4184
      Top             =   2160
      Visible         =   0   'False
      Width           =   3750
   End
   Begin VB.Image img1 
      Height          =   1215
      Left            =   240
      Picture         =   "frmFight1.frx":5BEB
      Top             =   2160
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.Label lbl2 
      Caption         =   "Inventory:"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lbl1 
      Caption         =   "Dragon Fight!!!!  Look in your inventory to see what you can use to fight the dragon.  You only get one weapon so choose wisely."
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
   End
End
Attribute VB_Name = "frmFight1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Katie Deeney & Elise Generex
'Create a Story
'Date Done: 10/10/2009
'In this form, it checks your inventory by how many letters are in it
'It allows the user to use what weapons are allowed and then brings them to the form
'that has the results on it

Private Sub cmdInventory_Click()
    If Len(Inventory) = 7 Then
        img3.Visible = True
        cmdSword.Visible = True
    ElseIf Len(Inventory) = 9 Then
        img4.Visible = True
        cmdKungFu.Visible = True
    ElseIf Len(Inventory) = 10 Then
        img2.Visible = True
        cmdsandwich.Visible = True
    ElseIf Len(Inventory) = 0 Then
        img1.Visible = True
        cmdToad.Visible = True
    ElseIf Len(Inventory) = 16 Then
        img4.Visible = True
        cmdKungFu.Visible = True
        img3.Visible = True
        cmdSword.Visible = True
    ElseIf Len(Inventory) = 17 Then
        img3.Visible = True
        cmdSword.Visible = True
        img2.Visible = True
        cmdsandwich.Visible = True
    ElseIf Len(Inventory) = 26 Then
        img4.Visible = True
        cmdKungFu.Visible = True
        img3.Visible = True
        cmdSword.Visible = True
        img2.Visible = True
        cmdsandwich.Visible = True
    ElseIf Len(Inventory) = 26 Then
        img4.Visible = True
        cmdKungFu.Visible = True
        img2.Visible = True
        cmdsandwich.Visible = True
    Else
        MsgBox "error", , "alert"
    End If
End Sub

Private Sub cmdKungFu_Click()
    frmFight1.Hide
    frmFightKungFu.Show
    img4.Visible = False
    cmdKungFu.Visible = False
    img3.Visible = False
    cmdSword.Visible = False
    img2.Visible = False
    cmdToad.Visible = False
    img1.Visible = False
    cmdsandwich.Visible = False
End Sub

Option Explicit

'Katie Deeney & Elise Generex
'Create a Story
'Date Done: 10/10/2009
'This is the users inventory, they have to pick from the objects
'they got from the caves to fight the dragon
Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdsandwich_Click()
    frmFight1.Hide
    MsgBox "The Dragon was very very hungry, so he ate the sandwich. In return, he let you have the princess. Your warts are healed! You live happily ever after!", , "Sandwich"
    MsgBox "This is where your story ends. Start Over.", , "Story Ends"
    Inventory = ""
    img4.Visible = False
    cmdKungFu.Visible = False
    img3.Visible = False
    cmdSword.Visible = False
    img2.Visible = False
    cmdToad.Visible = False
    img1.Visible = False
    cmdsandwich.Visible = False
    frmWelcome.Show
End Sub

Private Sub cmdSword_Click()
    frmFight1.Hide
    frmFightSword.Show
    img4.Visible = False
    cmdKungFu.Visible = False
    img3.Visible = False
    cmdSword.Visible = False
    img2.Visible = False
    cmdToad.Visible = False
    img1.Visible = False
    cmdsandwich.Visible = False
End Sub

Private Sub cmdToad_Click()
    frmFight1.Hide
    frmFightToad.Show
    img4.Visible = False
    cmdKungFu.Visible = False
    img3.Visible = False
    cmdSword.Visible = False
    img2.Visible = False
    cmdToad.Visible = False
    img1.Visible = False
    cmdsandwich.Visible = False
End Sub
