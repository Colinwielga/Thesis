VERSION 5.00
Begin VB.Form colors 
   BackColor       =   &H00C000C0&
   Caption         =   "Apparel Color"
   ClientHeight    =   6105
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11790
   FillColor       =   &H00FF00FF&
   ForeColor       =   &H00FF00FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6105
   ScaleWidth      =   11790
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton advance 
      Caption         =   "NEXT"
      Height          =   495
      Left            =   6120
      TabIndex        =   15
      Top             =   5160
      Width           =   1335
   End
   Begin VB.PictureBox picresults 
      Height          =   375
      Left            =   4680
      ScaleHeight     =   315
      ScaleWidth      =   1035
      TabIndex        =   14
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton selected 
      Caption         =   "COLOR SELECTED"
      Height          =   495
      Left            =   2760
      TabIndex        =   13
      Top             =   5160
      Width           =   1695
   End
   Begin VB.CommandButton navy 
      BackColor       =   &H00400000&
      Caption         =   "NAVY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8160
      MaskColor       =   &H00400000&
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3720
      Width           =   2295
   End
   Begin VB.CommandButton brown 
      BackColor       =   &H00000040&
      Caption         =   "BROWN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1080
      Width           =   2295
   End
   Begin VB.CommandButton orange 
      BackColor       =   &H000080FF&
      Caption         =   "ORANGE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8160
      MaskColor       =   &H000080FF&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2400
      Width           =   2295
   End
   Begin VB.CommandButton purple 
      BackColor       =   &H00800080&
      Caption         =   "PURPLE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3720
      Width           =   2295
   End
   Begin VB.CommandButton green 
      BackColor       =   &H0000C000&
      Caption         =   "GREEN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3720
      Width           =   2295
   End
   Begin VB.CommandButton gray 
      Caption         =   "GRAY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   720
      TabIndex        =   7
      Top             =   3720
      Width           =   2295
   End
   Begin VB.CommandButton black 
      BackColor       =   &H00000000&
      Caption         =   "BLACK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2400
      Width           =   2295
   End
   Begin VB.CommandButton yellow 
      BackColor       =   &H0000FFFF&
      Caption         =   "YELLOW"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1080
      Width           =   2295
   End
   Begin VB.CommandButton pink 
      BackColor       =   &H008080FF&
      Caption         =   "PINK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2400
      Width           =   2295
   End
   Begin VB.CommandButton red 
      BackColor       =   &H000000FF&
      Caption         =   "RED"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3240
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1080
      Width           =   2295
   End
   Begin VB.CommandButton white 
      BackColor       =   &H00FFFFFF&
      Caption         =   "WHITE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   720
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2400
      Width           =   2295
   End
   Begin VB.CommandButton blue 
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      Caption         =   "BLUE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   720
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C000C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Please Select A Color For Your Apparel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   7455
   End
End
Attribute VB_Name = "colors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : Screen Printing(Main1.vpb)
'Form Name : colors(Colors.frm)
'Author: Elizabeth Beesch
'Date Written: October 27, 2003
'Purpose of Form: To have the user select the color
        ' of the apparel print the color selected
'Option Explicit is a command to force
'the user to explicitly declare all variables
'before they can be used.
Option Explicit
Private Sub Form_Load() ' tells the program what to do when this form is loaded
selected.Enabled = False ' dosen't allow use of this button
advance.Enabled = False ' dosen't allow use of this button
End Sub ' end of operations for this sub
Private Sub black_Click() 'when this button is clicked on
finalcolor = "Black" 'declares variable as a color
selected.Enabled = True 'allows users to click this button
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
End Sub ' end of operations for this sub
Private Sub brown_Click() 'when this button is clicked on
finalcolor = "Brown" 'declares variable as a color
selected.Enabled = True 'allows users to click this button
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
End Sub ' end of operations for this sub
Private Sub blue_Click() 'when this button is clicked on
finalcolor = "Blue" 'declares variable as a color
selected.Enabled = True 'allows users to click this button
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
End Sub ' end of operations for this sub
Private Sub gray_Click() 'when this button is clicked on
finalcolor = "Gray" 'declares variable as a color
selected.Enabled = True 'allows users to click this button
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
End Sub ' end of operations for this sub
Private Sub green_Click() 'when this button is clicked on
finalcolor = "Green" 'declares variable as a color
selected.Enabled = True 'allows users to click this button
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
End Sub ' end of operations for this sub
Private Sub navy_Click() 'when this button is clicked on
finalcolor = "Navy" 'declares variable as a color
selected.Enabled = True 'allows users to click this button
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
End Sub ' end of operations for this sub
Private Sub orange_Click() 'when this button is clicked on
finalcolor = "Orange" 'declares variable as a color
selected.Enabled = True 'allows users to click this button
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
End Sub ' end of operations for this sub
Private Sub pink_Click() ' end of operations for this sub
finalcolor = "Pink" 'declares variable as a color
selected.Enabled = True 'allows users to click this button
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
End Sub ' end of operations for this sub
Private Sub purple_Click() 'when this button is clicked on
finalcolor = "Purple" 'declares variable as a color
selected.Enabled = True 'allows users to click this button
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
End Sub ' end of operations for this sub
Private Sub red_Click() 'when this button is clicked on
finalcolor = "Red" 'declares variable as a color
selected.Enabled = True 'allows users to click this button
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
End Sub ' end of operations for this sub
Private Sub white_Click() 'when this button is clicked on
finalcolor = "White" 'declares variable as a color
selected.Enabled = True 'allows users to click this button
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
End Sub ' end of operations for this sub
Private Sub yellow_Click() 'when this button is clicked on
finalcolor = "Yellow"
selected.Enabled = True 'allows users to click this button
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
End Sub ' end of operations for this sub
Private Sub selected_Click() 'when this button is clicked on
advance.Enabled = True ' allows use of this button
selected.Enabled = False ' dosen't allow use of button
picResults.Print finalcolor ' prints finalcolor
End Sub ' end of operations for this sub
Private Sub advance_Click() 'when this button is clicked on
colors.Hide ' hides current form
inkcolors.Show ' shows next from in sequence
End Sub ' end of operations for this sub
