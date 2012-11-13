VERSION 5.00
Begin VB.Form frmFight2 
   Caption         =   "Fight the Dragon!!"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8085
   LinkTopic       =   "Form1"
   ScaleHeight     =   5790
   ScaleWidth      =   8085
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdQuit 
      Caption         =   "Quit"
      Height          =   195
      Left            =   4080
      TabIndex        =   7
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton cmdStone 
      Caption         =   "Use Stone"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   3480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdInventory 
      Caption         =   "Look in Inventory"
      Height          =   615
      Left            =   2040
      TabIndex        =   4
      Top             =   1680
      Width           =   2655
   End
   Begin VB.CommandButton cmdSword 
      Caption         =   "Use Sword"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   5160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdsandwich 
      Caption         =   "Use Sandwich"
      Height          =   255
      Left            =   2400
      TabIndex        =   2
      Top             =   5160
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdKungFu 
      Caption         =   "Use Kung Fu Master"
      Height          =   255
      Left            =   5400
      TabIndex        =   1
      Top             =   5040
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Shape shp1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   975
      Left            =   0
      Shape           =   2  'Oval
      Top             =   2760
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1335
      Left            =   0
      Top             =   2640
      Visible         =   0   'False
      Width           =   1455
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
      TabIndex        =   5
      Top             =   2160
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Image img3 
      Height          =   1560
      Left            =   0
      Picture         =   "frmFight2.frx":0000
      Top             =   3960
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Image img2 
      Height          =   2880
      Left            =   1440
      Picture         =   "frmFight2.frx":0A94
      Top             =   2640
      Visible         =   0   'False
      Width           =   3750
   End
   Begin VB.Image img4 
      Height          =   4980
      Left            =   5160
      Picture         =   "frmFight2.frx":24FB
      Top             =   600
      Visible         =   0   'False
      Width           =   2535
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
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   5295
   End
End
Attribute VB_Name = "frmFight2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Katie Deeney & Elise Generex
'Create a Story
'Date Done: 10/10/2009
'This is the users inventory, they have to pick from the objects
'they got from the caves to fight the dragon
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
        Shape1.Visible = True
        shp1.Visible = True
        cmdStone.Visible = True
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
    End If
End Sub

Private Sub cmdKungFu_Click()
    frmFight2.Hide
    frmFightKungFu2.Show
    img4.Visible = False
    cmdKungFu.Visible = False
    img3.Visible = False
    cmdSword.Visible = False
    img2.Visible = False
    cmdStone.Visible = False
    shp1.Visible = False
    Shape1.Visible = False
    cmdsandwich.Visible = False
    
End Sub

Private Sub CmdQuit_Click()
    End
End Sub

Private Sub cmdsandwich_Click()
    frmFight2.Hide
    MsgBox "The Dragon was very very hungry, so he ate the sandwich. In return, he let you have the princess. You live happily ever after!", , "Sandwich"
    MsgBox "This is where your story ends. Start Over.", , "Story Ends"
    Inventory = ""
    img4.Visible = False
    cmdKungFu.Visible = False
    img3.Visible = False
    cmdSword.Visible = False
    img2.Visible = False
    cmdStone.Visible = False
    shp1.Visible = False
    Shape1.Visible = False
    cmdsandwich.Visible = False
    frmWelcome.Show
End Sub

Private Sub cmdStone_Click()
    frmFight2.Hide
    frmFightStone.Show
    img4.Visible = False
    cmdKungFu.Visible = False
    img3.Visible = False
    cmdSword.Visible = False
    img2.Visible = False
    cmdStone.Visible = False
    shp1.Visible = False
    Shape1.Visible = False
    cmdsandwich.Visible = False
End Sub

Private Sub cmdSword_Click()
    frmFight2.Hide
    frmFightSword2.Show
    img4.Visible = False
    cmdKungFu.Visible = False
    img3.Visible = False
    cmdSword.Visible = False
    img2.Visible = False
    cmdStone.Visible = False
    shp1.Visible = False
    Shape1.Visible = False
    cmdsandwich.Visible = False
End Sub
