VERSION 5.00
Begin VB.Form FrmSurf 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "SURFING"
   ClientHeight    =   11010
   ClientLeft      =   105
   ClientTop       =   -195
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmSurf.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CmdEnd 
      BackColor       =   &H00808000&
      Caption         =   "Exit Program"
      BeginProperty Font 
         Name            =   "Lucida Sans Typewriter"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9720
      Width           =   2055
   End
   Begin VB.CommandButton cmdboards 
      BackColor       =   &H00808000&
      Caption         =   "The Boards"
      BeginProperty Font 
         Name            =   "Lucida Sans Typewriter"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6960
      Width           =   2055
   End
   Begin VB.CommandButton Cmdpros 
      BackColor       =   &H00808000&
      Caption         =   "The Pros"
      BeginProperty Font 
         Name            =   "Lucida Sans Typewriter"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3720
      Width           =   2055
   End
   Begin VB.CommandButton cmdbasics 
      BackColor       =   &H00808000&
      Caption         =   "The Basics"
      BeginProperty Font 
         Name            =   "Lucida Sans Typewriter"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   720
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5400
      Width           =   2055
   End
   Begin VB.CommandButton CmdDest 
      BackColor       =   &H00808000&
      Caption         =   "The Destinations"
      BeginProperty Font 
         Name            =   "Lucida Sans Typewriter"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   720
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label lblDisplay 
      BackStyle       =   0  'Transparent
      Caption         =   $"FrmSurf.frx":11A54
      BeginProperty Font 
         Name            =   "Lucida Sans Typewriter"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2655
      Left            =   8040
      TabIndex        =   5
      Top             =   4320
      Width           =   7095
   End
End
Attribute VB_Name = "FrmSurf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : SurfProject (SurfingProject.vbp)
'Form Name : frmSurf (frmSurf.frm) 'this is the main form
'Author: Benjamin Luther
'Date : Monday October 30, 2005
'Purpose of the Project: 'To inform and educate the user about surfing.
                        'Users can view the different destinations,
                        'professional surfers, the basics about how to surf,
                        'and the best surfboard for them.\
'Purpose of the Form: this form acts as an introduction to the user,
                        'presenting them with all contents of the program.
                        'This form presents the user with links to other forms
                        'presenting them with more information about the topics.
Option Explicit 'Forces explicit declaration of all variables
Private Sub Close_Click()
    FrmEnd.Show 'shows the exit form
End Sub

Private Sub cmdbasics_Click()
    frmBasics.Show 'displays the basics form,
                    'informing users about the basic skills and necessities of surfing.
End Sub

Private Sub cmdboards_Click()
    Dim F As Integer 'declairs storage space for variable
    Do Until F = 2 'Loop is continued until F=2
        Level = InputBox("Please Enter Your Level:                                                      1 for Beginners                                                                  2 for Intermediate                                                              3 for Advanced", "Input Skill Level") 'Asks user to input their skill level
        If Level < 4 Then 'If the current value of Level is equal to one continue though the If stament
            FrmBoards.Show 'displays the Boards form
            F = 2 'sets f=2 to end loop
        Else 'If the entry listed is invalid it asks the user to try again
            MsgBox "Invalid Entry, Please Try Again", , "Invalid" 'displays pop-up that states to retry their entry, then the Input box reappears
        End If 'ends the if statment
    Loop 'loops back to Do Until F=2
End Sub
Private Sub CmdDest_Click()
    FrmDest.Show 'shows the destinations map to inform user about surfing places
End Sub

Private Sub Cmdend_Click()
    FrmEnd.Show 'displays the end form, asks the user if they are sure they want to exit
End Sub

Private Sub Cmdpros_Click()
    FrmPros.Show 'shows the pros form to inform user about the current proffesionals and legends
End Sub

