VERSION 5.00
Begin VB.Form frmAttributes 
   BackColor       =   &H00000000&
   Caption         =   "Character Attributes"
   ClientHeight    =   8115
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5370
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8115
   ScaleWidth      =   5370
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdShow 
      Caption         =   "Show"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   3
      Top             =   2640
      Width           =   2415
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1800
      TabIndex        =   1
      Top             =   7320
      Width           =   1695
   End
   Begin VB.PictureBox picAttributes 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3975
      Left            =   1080
      ScaleHeight     =   3915
      ScaleWidth      =   3075
      TabIndex        =   0
      Top             =   3240
      Width           =   3135
   End
   Begin VB.Label lblPet 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Pet:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   120
      Width           =   2175
   End
   Begin VB.Image img1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1620
      Left            =   1680
      Picture         =   "frmAttributes.frx":0000
      Stretch         =   -1  'True
      Top             =   720
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Image img2 
      BorderStyle     =   1  'Fixed Single
      Height          =   1635
      Left            =   1680
      Picture         =   "frmAttributes.frx":2EA8
      Stretch         =   -1  'True
      Top             =   720
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Image img4 
      BorderStyle     =   1  'Fixed Single
      Height          =   1695
      Left            =   1680
      Picture         =   "frmAttributes.frx":54BA
      Stretch         =   -1  'True
      Top             =   720
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Image img5 
      BorderStyle     =   1  'Fixed Single
      Height          =   1725
      Left            =   1680
      Picture         =   "frmAttributes.frx":7CF5
      Stretch         =   -1  'True
      Top             =   720
      Visible         =   0   'False
      Width           =   1845
   End
   Begin VB.Image img6 
      BorderStyle     =   1  'Fixed Single
      Height          =   1695
      Left            =   1680
      Picture         =   "frmAttributes.frx":A3CB
      Stretch         =   -1  'True
      Top             =   720
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Image img3 
      BorderStyle     =   1  'Fixed Single
      Height          =   1650
      Left            =   1680
      Picture         =   "frmAttributes.frx":CAE9
      Stretch         =   -1  'True
      Top             =   720
      Visible         =   0   'False
      Width           =   1815
   End
End
Attribute VB_Name = "frmAttributes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: RPGCraze
'Form name: frmAttributes
'Author: Justin Roth
'Date Written: Sunday, November 4th, 2007
'Objective of form: This form displays the character's attributes and details to the user.
        'The user can view their strength, intelligence, agility, cash, health, and other information on this form.
        
Option Explicit

Private Sub cmdBack_Click()
    frmAttributes.Hide  'Goes back to the Map form.
End Sub

Private Sub cmdShow_Click()
    picAttributes.Cls   'Clears the attributes picture box.
    picAttributes.Print "Name: "; N 'Displays the character's name in the attributes box.
    picAttributes.Print "Height: " & HghtFt; " '"; HghtIn; """" 'Displays the character's height in the attributes box.
    picAttributes.Print "Weight: " & Weight; " lbs" 'Displays the character's name in the attributes box.
    picAttributes.Print "Gender: " & Gender 'Displays the character's gender in the attributes box.
    picAttributes.Print "Age: " & Age & " yrs"  'Displays the character's age in the attributes box.
    
    'Sets up a layout for the data.
    picAttributes.Print "--------------------------------"
    picAttributes.Print "Character Attributes"
    picAttributes.Print "--------------------------------"
    
    picAttributes.Print "Strength: "; Tab(21); Strength; "pts." 'Displays the user's current strength.
    If Strength = 0 Then    'Tells the user where to get strength points if they don't have any.
        MsgBox "To gain strength, go to the Mystery Forest.", , "Strength!"
    End If
    
    Intelligence = Score
    picAttributes.Print "Intelligence: "; Tab(21); Intelligence; "pts." 'Displays the user's current intelligence.
    If Intelligence = 0 Then    'Tells the user where to get intelligence points if they don't have any.
        MsgBox "To gain intelligence, go to the Casino.", , "Intelligence"
    End If
    
    picAttributes.Print "Agility: "; Tab(21); Agility; "pts."   'Displays the user's current agility.
    If Agility = 0 Then 'Tells the user where to get agility points if they don't have any.
        MsgBox "To gain agility, go to the Hospital.", , "Agility!"
    End If
    
    'Sets up a layout for the data.
    picAttributes.Print "--------------------------------"
    picAttributes.Print "Character Details"
    picAttributes.Print "--------------------------------"
    
    picAttributes.Print "Health: "; Tab(15); MyHealth; "/100"   'Displays the user's current health.
    If MyHealth < 100 Then  'Tells the user where to get health if they are low.
        MsgBox "Your health is below 100, you should go to the Hospital.", , "Health!"
    End If
    
    picAttributes.Print "Cash: "; Tab(15); FormatCurrency(Cash) 'Displays the user's current cash funds.
    If Cash = 0 Then    'Tells the user where to get cash if they don't have any.
        MsgBox "To get cash, go to the Casino.", , "Cash!"
    End If
    
End Sub

Private Sub Form_Load()
    'Displays the pet (at the top of the attributes form) that the user chose.
    If Pet = 1 Then
        img1.Visible = True
    ElseIf Pet = 2 Then
        img2.Visible = True
    ElseIf Pet = 3 Then
        img3.Visible = True
    ElseIf Pet = 4 Then
        img4.Visible = True
    ElseIf Pet = 5 Then
        img5.Visible = True
    ElseIf Pet = 6 Then
        img6.Visible = True
    End If
    
End Sub
