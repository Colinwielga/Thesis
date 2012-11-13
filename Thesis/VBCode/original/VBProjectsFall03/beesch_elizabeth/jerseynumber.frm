VERSION 5.00
Begin VB.Form jerseynumber 
   BackColor       =   &H00C0C000&
   Caption         =   "Jersey Number"
   ClientHeight    =   6870
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11835
   FillColor       =   &H00FFFF00&
   ForeColor       =   &H00C0C000&
   LinkTopic       =   "Form2"
   ScaleHeight     =   6870
   ScaleWidth      =   11835
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton advance 
      BackColor       =   &H00FFFF80&
      Caption         =   "NEXT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8640
      MaskColor       =   &H00FFFF80&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3480
      UseMaskColor    =   -1  'True
      Width           =   1215
   End
   Begin VB.CommandButton selected 
      BackColor       =   &H00FFFF80&
      Caption         =   "NUMBER SELECTED"
      Height          =   495
      Left            =   8640
      MaskColor       =   &H00FFFF80&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1800
      Width           =   1215
   End
   Begin VB.PictureBox picresults 
      Height          =   615
      Left            =   7920
      ScaleHeight     =   555
      ScaleWidth      =   2715
      TabIndex        =   5
      Top             =   2520
      Width           =   2775
   End
   Begin VB.OptionButton optnonumber 
      BackColor       =   &H00C0C000&
      Caption         =   "NO NUMBER"
      Height          =   255
      Left            =   5400
      TabIndex        =   3
      Top             =   600
      Width           =   255
   End
   Begin VB.OptionButton optnumber 
      BackColor       =   &H00C0C000&
      Caption         =   "Number"
      Height          =   255
      Left            =   2040
      TabIndex        =   2
      Top             =   720
      Width           =   255
   End
   Begin VB.Label formlabel 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      Caption         =   "PLEASE SELECT NUMBER OR NO NUMBER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   1095
      Left            =   7560
      TabIndex        =   4
      Top             =   480
      Width           =   3495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      Caption         =   "NO NUMBER"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5040
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "NUMBER"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1800
      TabIndex        =   0
      Top             =   480
      Width           =   735
   End
   Begin VB.Image nonumber 
      Height          =   3840
      Left            =   4440
      Picture         =   "jerseynumber.frx":0000
      Top             =   1080
      Width           =   2955
   End
   Begin VB.Image number 
      Height          =   3600
      Left            =   840
      Picture         =   "jerseynumber.frx":1034
      Top             =   1200
      Width           =   2565
   End
End
Attribute VB_Name = "jerseynumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: Screen Printing(Main1.vpb)
'Form Name : jerseynumber(jerseynumber.frm)
'Author: Elizabeth Beesch
'Date Written: October 27, 2003
'Purpose of Form: To have the user choose weather or not they are going
                  ' to have a number on their apparel
Option Explicit 'Option Explicit is a command to force
'the user to explicitly declare all variables
'before they can be used.

Private Sub Form_Load() ' when the form loads the following will happen
    selected.Enabled = False 'disables the selected button
    advance.Enabled = False ' disables the advance button
End Sub 'ends the commands of the form load

Private Sub advance_Click() ' when the form loads the following will happen
    jerseynumber.Hide ' hides current form
    colors.Show ' shows next form in sequence
End Sub 'ends the commands of the button

Private Sub optnonumber_Click() 'when this button is clicked on the following will happen
    C = 0 'changes the value of c
    selected.Enabled = True ' enables the selected button
End Sub 'ends the commands of the button

Private Sub optnumber_Click() 'when this button is clicked on the following will happen
    C = 1.5 ' change the value of c
    selected.Enabled = True ' enables the selected button
End Sub 'ends the commands of the button

Private Sub selected_Click() 'when this button is clicked on the following will happen
    numbercost = C * numberofItems ' calculates the total number cost
    picresults.Print "Total Cost for Number(s):"; FormatCurrency(numbercost, 2)
    selected.Enabled = False ' disables the selected button
    advance.Enabled = True ' enables the advance button
End Sub 'ends the commands of the button


