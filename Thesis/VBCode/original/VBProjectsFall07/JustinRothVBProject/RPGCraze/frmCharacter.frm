VERSION 5.00
Begin VB.Form frmCharacter 
   BackColor       =   &H00000000&
   Caption         =   "Character Information"
   ClientHeight    =   7290
   ClientLeft      =   1140
   ClientTop       =   1530
   ClientWidth     =   9930
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
   LinkTopic       =   "Form1"
   Picture         =   "frmCharacter.frx":0000
   ScaleHeight     =   7290
   ScaleWidth      =   9930
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdResetPets 
      Caption         =   "Reset pets"
      Height          =   255
      Left            =   7080
      TabIndex        =   6
      Top             =   6960
      Width           =   1335
   End
   Begin VB.CommandButton cmdQuest 
      Caption         =   "Start your Quest!"
      Height          =   735
      Left            =   1320
      TabIndex        =   5
      Top             =   5280
      Width           =   2415
   End
   Begin VB.PictureBox picInfo 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
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
      Height          =   2775
      Left            =   840
      ScaleHeight     =   2715
      ScaleWidth      =   3195
      TabIndex        =   3
      Top             =   2280
      Width           =   3255
   End
   Begin VB.PictureBox picName 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   960
      ScaleHeight     =   435
      ScaleWidth      =   2955
      TabIndex        =   2
      Top             =   1200
      Width           =   3015
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   6240
      Width           =   1575
   End
   Begin VB.CommandButton cmdInfo 
      BackColor       =   &H00000000&
      Caption         =   "Display your character information!"
      Height          =   735
      Left            =   960
      MaskColor       =   &H00000000&
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label lblPet 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Choose a pet by clicking on one :"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   5760
      TabIndex        =   4
      Top             =   240
      Width           =   3975
   End
   Begin VB.Image img1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1860
      Left            =   5400
      Picture         =   "frmCharacter.frx":2AF2E
      Stretch         =   -1  'True
      Top             =   720
      Width           =   2235
   End
   Begin VB.Image img2 
      BorderStyle     =   1  'Fixed Single
      Height          =   1875
      Left            =   7800
      Picture         =   "frmCharacter.frx":2DDD6
      Stretch         =   -1  'True
      Top             =   720
      Width           =   2055
   End
   Begin VB.Image img4 
      BorderStyle     =   1  'Fixed Single
      Height          =   2055
      Left            =   7800
      Picture         =   "frmCharacter.frx":303E8
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Image img5 
      BorderStyle     =   1  'Fixed Single
      Height          =   1965
      Left            =   5400
      Picture         =   "frmCharacter.frx":32C23
      Stretch         =   -1  'True
      Top             =   4920
      Width           =   2205
   End
   Begin VB.Image img6 
      BorderStyle     =   1  'Fixed Single
      Height          =   1935
      Left            =   7800
      Picture         =   "frmCharacter.frx":352F9
      Stretch         =   -1  'True
      Top             =   4920
      Width           =   2055
   End
   Begin VB.Image img3 
      BorderStyle     =   1  'Fixed Single
      Height          =   2010
      Left            =   5400
      Picture         =   "frmCharacter.frx":37A17
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Line ln1 
      BorderColor     =   &H00000000&
      X1              =   5040
      X2              =   5040
      Y1              =   240
      Y2              =   6960
   End
End
Attribute VB_Name = "frmCharacter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: RPGCraze
'Form name: frmCharacter
'Author: Justin Roth
'Date Written: Sunday, November 4th, 2007
'Objective of form: This form is where the user can display their character information.
        'This form also allows the user to pick a pet for the progress of the game.

Option Explicit

Private Sub cmdInfo_Click()
    picInfo.Cls 'Clears the Character Information Box when the Info button is pushed.
    picName.Cls 'Clears the Character Name Box when the Info button is pushed.
    
    MsgBox "Make sure you select a pet before you go!", , "Select a pet!"   'Reminds the user to select a pet before they continue on.
    
    'This converts height (ft/inches) in decimal form to the standard 0'0" form.
    HghtFt = Int(Hght / 12) 'Splices off the numbers to the right of the decimal to leave a single digit.
    HghtIn = Hght - (HghtFt * 12)   'Converts the digits to the right of the decimal into inches out of 12 feet.
    
    picName.Print "Welcome, " & N; "!"  'Displays the character's name in a picture box.
    picInfo.Print "Character Information"   'Sets a layout for the data.
    picInfo.Print "----------------------"  'Sets a layout for the data.
    picInfo.Print "Height: " & HghtFt; " '"; HghtIn; """"   'Displays the character's height in the 0'0" format in a picture box.
    picInfo.Print "Weight: " & Weight; " lbs"   'Displays the character's weight in lbs.
    picInfo.Print "Gender: " & Gender   'Displays the character's gender in a picture box.
    picInfo.Print "Age: " & Age & " yrs"    'Displays the character's age in a picture box.
    
End Sub

Private Sub cmdAttributes_Click()
    frmAttributes.Show  'Shows the attributes for the user.
End Sub

Private Sub cmdQuest_Click()
    frmMap.Show 'Allows the user to navigate to the Map form to begin their quest.
    frmCharacter.Hide   'Hides the character form so the user can move on.
End Sub

Private Sub cmdQuit_Click()
    End 'Quits the program.
End Sub

Private Sub cmdResetPets_Click()    'Resets all of the pets so the user can reselect one if needed.
    img1.Visible = True
    img2.Visible = True
    img3.Visible = True
    img4.Visible = True
    img5.Visible = True
    img6.Visible = True
End Sub

Private Sub img1_Click()    'Hides all other pet options except for Pet 1(img1).
    Pet = 1 'Assigns the first pet a public integer for use in the attributes form.
    img1.Visible = True
    img2.Visible = False
    img3.Visible = False
    img4.Visible = False
    img5.Visible = False
    img6.Visible = False
End Sub

Private Sub img2_Click()    'Hides all other pet options except for Pet 2(img2).
    Pet = 2 'Assigns the second pet a public integer for use in the attributes form.
    img1.Visible = False
    img2.Visible = True
    img3.Visible = False
    img4.Visible = False
    img5.Visible = False
    img6.Visible = False
End Sub

Private Sub img3_Click()    'Hides all other pet options except for Pet 3(img3).
    Pet = 3 'Assigns the third pet a public integer for use in the attributes form.
    img1.Visible = False
    img2.Visible = False
    img3.Visible = True
    img4.Visible = False
    img5.Visible = False
    img6.Visible = False
End Sub

Private Sub img4_Click()    'Hides all other pet options except for Pet 4(img4).
    Pet = 4 'Assigns the fourth pet a public integer for use in the attributes form.
    img1.Visible = False
    img2.Visible = False
    img3.Visible = False
    img4.Visible = True
    img5.Visible = False
    img6.Visible = False
End Sub

Private Sub img5_Click()    'Hides all other pet options except for Pet 5(img5).
    Pet = 5 'Assigns the fifth pet a public integer for use in the attributes form.
    img1.Visible = False
    img2.Visible = False
    img3.Visible = False
    img4.Visible = False
    img5.Visible = True
    img6.Visible = False
End Sub

Private Sub img6_Click()    'Hides all other pet options except for Pet 6(img6).
    Pet = 6 'Assigns the sixth pet a public integer for use in the attributes form.
    img1.Visible = False
    img2.Visible = False
    img3.Visible = False
    img4.Visible = False
    img5.Visible = False
    img6.Visible = True
End Sub
