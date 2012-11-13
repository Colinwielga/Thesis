VERSION 5.00
Begin VB.Form jerseyname 
   BackColor       =   &H00800080&
   Caption         =   "jerseyname"
   ClientHeight    =   6900
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12315
   FillColor       =   &H00C000C0&
   ForeColor       =   &H00FF00FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6900
   ScaleWidth      =   12315
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      Height          =   615
      Left            =   7800
      ScaleHeight     =   555
      ScaleWidth      =   2595
      TabIndex        =   7
      Top             =   3480
      Width           =   2655
   End
   Begin VB.CommandButton selected 
      Caption         =   "Name Selected"
      Height          =   495
      Left            =   8400
      TabIndex        =   6
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton advance 
      BackColor       =   &H00FF80FF&
      Caption         =   "NEXT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8520
      MaskColor       =   &H00FF00FF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4560
      UseMaskColor    =   -1  'True
      Width           =   1215
   End
   Begin VB.OptionButton optnoname 
      BackColor       =   &H00800080&
      Caption         =   "Option2"
      Height          =   255
      Left            =   5760
      TabIndex        =   1
      Top             =   4920
      Width           =   255
   End
   Begin VB.OptionButton optname 
      BackColor       =   &H00800080&
      Caption         =   "Option1"
      Height          =   255
      Left            =   2160
      TabIndex        =   0
      Top             =   4680
      Width           =   255
   End
   Begin VB.Label please 
      Alignment       =   2  'Center
      BackColor       =   &H00800080&
      BackStyle       =   0  'Transparent
      Caption         =   "Please Select Name or No Name"
      BeginProperty Font 
         Name            =   "TLEastEurope"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   2055
      Left            =   7920
      TabIndex        =   4
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label noname 
      Alignment       =   2  'Center
      BackColor       =   &H00800080&
      BackStyle       =   0  'Transparent
      Caption         =   "NO NAME"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   5520
      TabIndex        =   3
      Top             =   480
      Width           =   855
   End
   Begin VB.Label name 
      Alignment       =   2  'Center
      BackColor       =   &H00800080&
      BackStyle       =   0  'Transparent
      Caption         =   "NAME"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1920
      TabIndex        =   2
      Top             =   480
      Width           =   495
   End
   Begin VB.Image picname 
      Height          =   3600
      Left            =   1080
      Picture         =   "jerseyname.frx":0000
      Top             =   960
      Width           =   2565
   End
   Begin VB.Image picnoname 
      Height          =   3840
      Left            =   4320
      Picture         =   "jerseyname.frx":15D8
      Top             =   960
      Width           =   2955
   End
End
Attribute VB_Name = "jerseyname"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: Screen Printing(Main1.vpb)
'Form Name : jerseyname(jerseyname.frm)
'Author: Elizabeth Beesch
'Date Written: October 27, 2003
'Purpose of Form: To have the user choose weather or not they are going
                  ' to have a name on their apparel
Option Explicit 'Option Explicit is a command to force
'the user to explicitly declare all variables
'before they can be used.

Private Sub Form_Load() ' when the form loads the following will happen
advance.Enabled = False ' disables the advance button
End Sub 'ends the commands of the form load

Private Sub optname_Click() 'when this button is clicked on the following will happen
B = 2 ' changes the value of B
End Sub 'ends the commands of the button

Private Sub optnoname_Click() 'when this button is clicked on the following will happen
B = 0 ' changes the value of B
End Sub 'ends the commands of the button

Private Sub advance_Click() 'when this button is clicked on the following will happen
jerseyname.Hide 'hides current form
jerseynumber.Show ' shows next form in sequence
End Sub ''ends the commands of the button

Private Sub selected_Click() 'when this button is clicked on the following will happen
namecost = B * numberofItems ' calculates the total name cost
picResults.Print "The Total Name Cost:"; FormatCurrency(namecost, 2)
selected.Enabled = False ' disables the selected button
advance.Enabled = True ' enables the advance button
End Sub 'ends the commands of the button


