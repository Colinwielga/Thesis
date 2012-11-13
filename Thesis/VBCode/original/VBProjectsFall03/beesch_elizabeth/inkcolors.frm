VERSION 5.00
Begin VB.Form inkcolors 
   BackColor       =   &H00C00000&
   Caption         =   "Ink Colors"
   ClientHeight    =   6405
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11115
   LinkTopic       =   "Form1"
   ScaleHeight     =   6405
   ScaleWidth      =   11115
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton nextbutton 
      Caption         =   "NEXT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3720
      Width           =   1335
   End
   Begin VB.PictureBox picResults 
      Height          =   975
      Left            =   2880
      ScaleHeight     =   915
      ScaleWidth      =   1275
      TabIndex        =   14
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CheckBox navy 
      BackColor       =   &H00C00000&
      Caption         =   "NAVY"
      Height          =   195
      Left            =   1200
      TabIndex        =   12
      Top             =   3840
      Width           =   855
   End
   Begin VB.CheckBox orange 
      BackColor       =   &H00C00000&
      Caption         =   "ORANGE"
      Height          =   195
      Left            =   600
      TabIndex        =   11
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CheckBox brown 
      BackColor       =   &H00C00000&
      Caption         =   "BROWN"
      Height          =   195
      Left            =   1200
      TabIndex        =   10
      Top             =   3120
      Width           =   975
   End
   Begin VB.CheckBox purple 
      BackColor       =   &H00C00000&
      Caption         =   "PURPLE"
      Height          =   195
      Left            =   600
      TabIndex        =   9
      Top             =   2760
      Width           =   975
   End
   Begin VB.CheckBox black 
      BackColor       =   &H00C00000&
      Caption         =   "BLACK"
      Height          =   195
      Left            =   1200
      TabIndex        =   8
      Top             =   2400
      Width           =   855
   End
   Begin VB.CheckBox yellow 
      BackColor       =   &H00C00000&
      Caption         =   "YELLOW"
      Height          =   195
      Left            =   600
      TabIndex        =   7
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CheckBox green 
      BackColor       =   &H00C00000&
      Caption         =   "GREEN"
      Height          =   255
      Left            =   1200
      TabIndex        =   6
      Top             =   1800
      Width           =   975
   End
   Begin VB.CheckBox pink 
      BackColor       =   &H00C00000&
      Caption         =   "PINK"
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   1440
      Width           =   735
   End
   Begin VB.CheckBox red 
      BackColor       =   &H00C00000&
      Caption         =   "RED"
      Height          =   195
      Left            =   1200
      TabIndex        =   4
      Top             =   1200
      Width           =   735
   End
   Begin VB.CheckBox gray 
      BackColor       =   &H00C00000&
      Caption         =   "GRAY"
      Height          =   195
      Left            =   600
      TabIndex        =   3
      Top             =   960
      Width           =   855
   End
   Begin VB.CheckBox white 
      BackColor       =   &H00C00000&
      Caption         =   "WHITE"
      Height          =   195
      Left            =   1200
      TabIndex        =   2
      Top             =   720
      Width           =   855
   End
   Begin VB.CheckBox blue 
      BackColor       =   &H00C00000&
      Caption         =   "BLUE"
      Height          =   195
      Left            =   600
      TabIndex        =   1
      Top             =   360
      Width           =   735
   End
   Begin VB.CommandButton selectedbutton 
      Caption         =   "Color Selected"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   0
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      Caption         =   "PLEASE SELECET COLOR FOR INK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   2640
      TabIndex        =   13
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "inkcolors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: Screen Printing(Main1.vpb)
'Form Name : inkcolors(inkcolors.frm)
'Author: Elizabeth Beesch
'Date Written: October 27, 2003
'Purpose of Form: To have the user choose a color for the ink
    'the user to explicitly declare all variables
    'before they can be used.
Option Explicit 'Option Explicit is a command to force
'the user to explicitly declare all variables
'before they can be used.

Private Sub Form_Load() 'when the form loads the following will happen
    nextbutton.Enabled = False 'disables the next button
    selectedbutton.Enabled = False 'disables the selected button
    End Sub

Private Sub black_Click() ' when the user clicks on this button the following will happen
    picResults.Print "Black"  'prints color of choice
    inkcolor = "Black" 'changes inkcolor
    selectedbutton.Enabled = True 'allows user to click on the selecetd button
    brown.Enabled = False ' enables the button
    blue.Enabled = False ' enables the button
    gray.Enabled = False ' enables the button
    red.Enabled = False ' enables the button
    pink.Enabled = False ' enables the button
    green.Enabled = False ' enables the button
    yellow.Enabled = False ' enables the button
    black.Enabled = False ' enables the button
    purple.Enabled = False ' enables the button
    brown.Enabled = False ' enables the button
    orange.Enabled = False ' enables the button
    navy.Enabled = False ' enables the button
white.Enabled = False 'enables the button
End Sub ' ends the commands of the button

Private Sub blue_Click() ' when the user clicks on this button the following will happen
    picResults.Print "Blue" 'prints color of choice
    inkcolor = "Blue" 'changes inkcolor
    selectedbutton.Enabled = True 'allows user to click on the selecetd button
    brown.Enabled = False ' enables the button
    blue.Enabled = False ' enables the button
    gray.Enabled = False ' enables the button
    red.Enabled = False ' enables the button
    pink.Enabled = False ' enables the button
    green.Enabled = False ' enables the button
    yellow.Enabled = False ' enables the button
    black.Enabled = False ' enables the button
    purple.Enabled = False ' enables the button
    brown.Enabled = False ' enables the button
    orange.Enabled = False ' enables the button
    navy.Enabled = False ' enables the button
    white.Enabled = False 'enables the button
End Sub ' ends the commands of the button

Private Sub brown_Click() ' when the user clicks on this button the following will happen
    picResults.Print "Brown" 'prints color of choice
    inkcolor = "Brown" 'changes inkcolor
    selectedbutton.Enabled = True 'allows user to click on the selecetd button
    brown.Enabled = False ' enables the button
    blue.Enabled = False ' enables the button
    gray.Enabled = False ' enables the button
    red.Enabled = False ' enables the button
    pink.Enabled = False ' enables the button
    green.Enabled = False ' enables the button
    yellow.Enabled = False ' enables the button
    black.Enabled = False ' enables the button
    purple.Enabled = False ' enables the button
    brown.Enabled = False ' enables the button
    orange.Enabled = False ' enables the button
    navy.Enabled = False ' enables the button
    white.Enabled = False 'enables the button
End Sub ' ends the commands of the button

Private Sub gray_Click() ' when the user clicks on this button the following will happen
    picResults.Print "Gray" 'prints color of choice
    inkcolor = "Gray" 'changes inkcolor
    selectedbutton.Enabled = True 'allows user to click on the selecetd button
    brown.Enabled = False ' enables the button
    blue.Enabled = False ' enables the button
    gray.Enabled = False ' enables the button
    red.Enabled = False ' enables the button
    pink.Enabled = False ' enables the button
    green.Enabled = False ' enables the button
    yellow.Enabled = False ' enables the button
    black.Enabled = False ' enables the button
    purple.Enabled = False ' enables the button
    brown.Enabled = False ' enables the button
    orange.Enabled = False ' enables the button
    navy.Enabled = False ' enables the button
    white.Enabled = False 'enables the button
End Sub ' ends the commands of the button

Private Sub green_Click() ' when the user clicks on this button the following will happen
    picResults.Print "Green" 'prints color of choice
    inkcolor = "Green" 'changes inkcolor
    selectedbutton.Enabled = True 'allows user to click on the selecetd button
    brown.Enabled = False ' enables the button
    blue.Enabled = False ' enables the button
    gray.Enabled = False ' enables the button
    red.Enabled = False ' enables the button
    pink.Enabled = False ' enables the button
    green.Enabled = False ' enables the button
    yellow.Enabled = False ' enables the button
    black.Enabled = False ' enables the button
    purple.Enabled = False ' enables the button
    brown.Enabled = False ' enables the button
    orange.Enabled = False ' enables the button
    navy.Enabled = False ' enables the button
    white.Enabled = False 'enables the button
End Sub ' ends the commands of the button

Private Sub navy_Click() ' when the user clicks on this button the following will happen
    picResults.Print "Navy" 'prints color of choice
    inkcolor = "Navy" 'changes inkcolor
    selectedbutton.Enabled = True 'allows user to click on the selecetd button
    brown.Enabled = False ' enables the button
    blue.Enabled = False ' enables the button
    gray.Enabled = False ' enables the button
    red.Enabled = False ' enables the button
    pink.Enabled = False ' enables the button
    green.Enabled = False ' enables the button
    yellow.Enabled = False ' enables the button
    lack.Enabled = False ' enables the button
    purple.Enabled = False ' enables the button
    brown.Enabled = False ' enables the button
    orange.Enabled = False ' enables the button
    navy.Enabled = False ' enables the button
    white.Enabled = False 'enables the button
End Sub ' ends the commands of the button

Private Sub orange_Click() ' when the user clicks on this button the following will happen
    picResults.Print "Orange" 'prints color of choice
    inkcolor = "Orange" 'changes inkcolor
    selectedbutton.Enabled = True 'allows user to click on the selecetd button
    brown.Enabled = False ' enables the button
    blue.Enabled = False ' enables the button
    gray.Enabled = False ' enables the button
    red.Enabled = False ' enables the button
    pink.Enabled = False ' enables the button
    green.Enabled = False ' enables the button
    yellow.Enabled = False ' enables the button
    black.Enabled = False ' enables the button
    purple.Enabled = False ' enables the button
    brown.Enabled = False ' enables the button
    orange.Enabled = False ' enables the button
    navy.Enabled = False ' enables the button
    white.Enabled = False 'enables the button
End Sub ' ends the commands of the button

Private Sub pink_Click() ' when the user clicks on this button the following will happen
    picResults.Print "Pink" 'prints color of choice
    inkcolor = "Pink" 'changes inkcolor
    selectedbutton.Enabled = True 'allows user to click on the selecetd button
    brown.Enabled = False ' enables the button
    blue.Enabled = False ' enables the button
    gray.Enabled = False ' enables the button
    red.Enabled = False ' enables the button
    pink.Enabled = False ' enables the button
    green.Enabled = False ' enables the button
    yellow.Enabled = False ' enables the button
    black.Enabled = False ' enables the button
    purple.Enabled = False ' enables the button
    brown.Enabled = False ' enables the button
    orange.Enabled = False ' enables the button
    navy.Enabled = False ' enables the button
    white.Enabled = False 'enables the button
End Sub ' ends the commands of the button

Private Sub purple_Click() ' when the user clicks on this button the following will happen
    picResults.Print "Purple" 'prints color of choice
    inkcolor = "Purple" 'changes inkcolor
    selectedbutton.Enabled = True 'allows user to click on the selecetd button
    brown.Enabled = False ' enables the button
    blue.Enabled = False ' enables the button
    gray.Enabled = False ' enables the button
    red.Enabled = False ' enables the button
    pink.Enabled = False ' enables the button
    green.Enabled = False ' enables the button
    yellow.Enabled = False ' enables the button
    black.Enabled = False ' enables the button
    purple.Enabled = False ' enables the button
    brown.Enabled = False ' enables the button
    orange.Enabled = False ' enables the button
    navy.Enabled = False ' enables the button
    white.Enabled = False 'enables the button
End Sub ' ends the commands of the button

Private Sub red_Click() ' when the user clicks on this button the following will happen
    picResults.Print "Red" 'prints color of choice
    inkcolor = "Red" 'changes inkcolor
    selectedbutton.Enabled = True 'allows user to click on the selecetd button
    brown.Enabled = False ' enables the button
    blue.Enabled = False ' enables the button
    gray.Enabled = False ' enables the button
    red.Enabled = False ' enables the button
    pink.Enabled = False ' enables the button
    green.Enabled = False ' enables the button
    yellow.Enabled = False ' enables the button
    black.Enabled = False ' enables the button
    purple.Enabled = False ' enables the button
    brown.Enabled = False ' enables the button
    orange.Enabled = False ' enables the button
    navy.Enabled = False ' enables the button
    white.Enabled = False 'enables the button
End Sub ' ends the commands of the button

Private Sub yellow_Click() ' when the user clicks on this button the following will happen
    picResults.Print "Yellow" 'prints color of choice
    inkcolor = "Yellow" 'changes inkcolor
    selectedbutton.Enabled = True 'allows user to click on the selecetd button
    brown.Enabled = False ' enables the button
    blue.Enabled = False ' enables the button
    gray.Enabled = False ' enables the button
    red.Enabled = False ' enables the button
    pink.Enabled = False ' enables the button
    green.Enabled = False ' enables the button
    yellow.Enabled = False ' enables the button
    black.Enabled = False ' enables the button
    purple.Enabled = False ' enables the button
    brown.Enabled = False ' enables the button
    orange.Enabled = False ' enables the button
    navy.Enabled = False ' enables the button
    white.Enabled = False 'enables the button
End Sub ' ends the commands of the button

Private Sub nextbutton_Click() 'when the user clicks on this button the following will happen
    Locations.Show 'shows the locations form
    inkcolors.Hide 'hides the current form inkcolros
End Sub ' ends the commands of the button

Private Sub selectedbutton_Click() ' when the user clicks on this button the following will happen
    selectedbutton.Enabled = False 'disables the selected button
    nextbutton.Enabled = True ' enables the next button
End Sub ' ends the commands of the button
