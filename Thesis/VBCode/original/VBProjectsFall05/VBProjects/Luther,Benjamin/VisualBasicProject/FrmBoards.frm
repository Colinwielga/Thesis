VERSION 5.00
Begin VB.Form FrmBoards 
   BackColor       =   &H00FFFFFF&
   Caption         =   "The BOARDS - Beginners"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15195
   BeginProperty Font 
      Name            =   "Myriad Web Pro"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00800000&
   LinkTopic       =   "Form1"
   Picture         =   "FrmBoards.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   15195
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmsize 
      BackColor       =   &H000000C0&
      Caption         =   "Find Size"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13440
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   6240
      Width           =   1095
   End
   Begin VB.PictureBox picsize 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   375
      Left            =   12360
      ScaleHeight     =   315
      ScaleWidth      =   2235
      TabIndex        =   22
      Top             =   7080
      Width           =   2295
   End
   Begin VB.TextBox txtenter 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12480
      TabIndex        =   20
      Top             =   6240
      Width           =   735
   End
   Begin VB.PictureBox picguns 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5535
      Left            =   2640
      Picture         =   "FrmBoards.frx":3DA4F
      ScaleHeight     =   5475
      ScaleWidth      =   1275
      TabIndex        =   14
      Top             =   4560
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox piclong 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   2640
      Picture         =   "FrmBoards.frx":4291A
      ScaleHeight     =   5235
      ScaleWidth      =   1275
      TabIndex        =   12
      Top             =   4680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdguns 
      BackColor       =   &H0000FFFF&
      Caption         =   "Guns"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12600
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3480
      Width           =   2295
   End
   Begin VB.CommandButton CmdFish 
      BackColor       =   &H000000C0&
      Caption         =   "Fish Boards"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3480
      Width           =   2295
   End
   Begin VB.CommandButton Cmdclose 
      BackColor       =   &H000000C0&
      Caption         =   "Close Screen"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   9840
      Width           =   1695
   End
   Begin VB.PictureBox picdisplay 
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "Lucida Sans Typewriter"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2655
      Left            =   4440
      ScaleHeight     =   2595
      ScaleWidth      =   7035
      TabIndex        =   3
      Top             =   600
      Width           =   7095
   End
   Begin VB.CommandButton Cmdsoftboards 
      BackColor       =   &H0000FFFF&
      Caption         =   "Softboards"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      MaskColor       =   &H000000C0&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3480
      Width           =   2295
   End
   Begin VB.CommandButton cmdshortboard 
      BackColor       =   &H0000FFFF&
      Caption         =   "Shortboards"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3480
      Width           =   2295
   End
   Begin VB.CommandButton cmdlongboard 
      BackColor       =   &H000000C0&
      Caption         =   "Longboards"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3480
      Width           =   2295
   End
   Begin VB.PictureBox picsoft 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6855
      Left            =   2400
      Picture         =   "FrmBoards.frx":47AF0
      ScaleHeight     =   6795
      ScaleWidth      =   1635
      TabIndex        =   10
      Top             =   3720
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.PictureBox picfish 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7335
      Left            =   2040
      Picture         =   "FrmBoards.frx":4DEE4
      ScaleHeight     =   7275
      ScaleWidth      =   2355
      TabIndex        =   11
      Top             =   3600
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.PictureBox picshort 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6975
      Left            =   1320
      Picture         =   "FrmBoards.frx":4F445
      ScaleHeight     =   6915
      ScaleWidth      =   3795
      TabIndex        =   13
      Top             =   3960
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Label Lblsize 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      Caption         =   "Enter the size of the wave in feet to find out what type of wave it is."
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   12480
      TabIndex        =   21
      Top             =   5280
      Width           =   2055
   End
   Begin VB.Label lblfish 
      BackColor       =   &H00400000&
      Caption         =   $"FrmBoards.frx":51A0F
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   5520
      TabIndex        =   19
      Top             =   5760
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.Label lblguns 
      BackColor       =   &H00400000&
      Caption         =   $"FrmBoards.frx":51B7C
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   5520
      TabIndex        =   18
      Top             =   5640
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.Label lblsoft 
      BackColor       =   &H00400000&
      Caption         =   $"FrmBoards.frx":51CF5
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2055
      Left            =   5520
      TabIndex        =   17
      Top             =   5520
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.Label lbllong 
      BackColor       =   &H00400000&
      Caption         =   $"FrmBoards.frx":51E7F
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   5400
      TabIndex        =   16
      Top             =   5040
      Visible         =   0   'False
      Width           =   5655
   End
   Begin VB.Label lblshort 
      BackColor       =   &H00400000&
      Caption         =   $"FrmBoards.frx":51FF9
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2415
      Left            =   5400
      TabIndex        =   15
      Top             =   5760
      Visible         =   0   'False
      Width           =   5775
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      Caption         =   "Your Level Information/ Recommendations "
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   4680
      TabIndex        =   6
      Top             =   240
      Width           =   6615
   End
   Begin VB.Label lbltitle2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " BOARDS"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   42
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   -1  'True
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   840
      Left            =   600
      TabIndex        =   5
      Top             =   960
      Width           =   3480
   End
   Begin VB.Label lbltitile 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "THE"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   -1  'True
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   720
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "FrmBoards"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: SurfProject (SurfingProject.vbp)
'Form Name: frmDest (frmBoards.frm)
'Author: Benjamin Luther
'Purpose of Form: the purpose of this form is to show users
                    'the type of board to use depending on their
                    'level in which they specify. The user interacts
                    'to find which is best for them
Option Explicit 'Forces explicit declaration of all variables.
Private Sub cmsize_Click() 'this button is designed to find the size of the wave that the user inputs
    Dim Size As Integer 'declares storage space for variable size
    picsize.Cls 'clears the size picture box
    Let Size = txtenter.Text 'sets the variable size equal to the txtbox enter
    If Size < 4 Then 'if size is less than 4 then continue
        picsize.Print Size; " feet: Small Wave" 'prints the wave size
    ElseIf Size < 10 Then 'if size is less than 10 continue
        picsize.Print Size; " feet: Medium Wave" 'prints the wave size
    ElseIf Size >= 10 Then 'if the size is greater than or equal to 10 then continue
        picsize.Print Size; " feet: Large Wave" 'prints the wave size
    End If 'ends the IF statement
End Sub

Private Sub Form_Activate() 'when this form is activated this information will be shown
    picdisplay.Cls 'clears anything that might be in the display picture box
    If Level = 1 Then 'if the variable level is equal to 1 then print the following in the display picture box
        picdisplay.Print Tab(22); "-Beginner Level-" ' tabs out 22 spaces then prints "-Beginner Level-" in the first line of the picture box
        picdisplay.Print "Your on your way to becoming a surfer,"
        picdisplay.Print "Gennerally beginners start with a softboard:" 'this line and the following lines until the If statement will be displayed in the picure box
        picdisplay.Print "A softboard delivers much more stability to the beginner," 'prints text in display picture box
        picdisplay.Print "and performs much slower allowing the user to master" 'prints text in display picture box
        picdisplay.Print "the basics of paddling, standing, and riding a wave." 'prints text in display picture box
        picdisplay.Print "Another option for the beginner is to start with a " 'prints text in display picture box
        picdisplay.Print "longboard, this also can provide the user with more" 'prints text in display picture box
        picdisplay.Print "stability for getting into the waves." 'prints text in display picture box
        picdisplay.Print 'prints a blank line in the display picture box
        picdisplay.Print "****Click the buttons for information on these boards!****" 'prints text in display picture box and informs user of the buttons below
    End If 'ends the If stament
    If Level = 2 Then 'if the variable level is equal to 2 then print the following in the display picture box
        picdisplay.Print Tab(20); "-Intermediate Level-" ' tabs out 20 spaces then prints "-Intermediate Level-" in the first line of the picture box
        picdisplay.Print "Alright, so your surfer with a little more skill that" 'prints text in display picture box
        picdisplay.Print "requires a more agile and lightweight board. Gennerally," 'prints text in display picture box
        picdisplay.Print "surfers of your level surf medium height waves with" 'prints text in display picture box
        picdisplay.Print "decent tow. Of the types of boards for medium surf " 'prints text in display picture box
        picdisplay.Print "Shortboards, and the Fishes work the best. Also," 'prints text in display picture box
        picdisplay.Print "for riding larger waves are the Longboards. They provide"  'prints text in display picture box
        picdisplay.Print "more stability and ease when learning to surf larger waves." 'prints text in display picture box
        picdisplay.Print 'prints a blank line in the display picture box
        picdisplay.Print "****Click the buttons for information on these boards!****" 'prints text in display picture box and informs user of the buttons below
    End If 'ends the If stament
    If Level = 3 Then 'if the variable level is equal to 3 then print the following in the display picture box
        picdisplay.Print Tab(22); "-Advanced Level-" ' tabs out 22 spaces then prints "-Advanced Level-" in the first line of the picture box
        picdisplay.Print "You among the top surfers and know all the fundamentals" 'prints text in display picture box
        picdisplay.Print "and advanced skills of surfing to surf the big waves." 'prints text in display picture box
        picdisplay.Print "Advanced surfers are able to use most boards to their" 'prints text in display picture box
        picdisplay.Print "likings, depending on their choice of wave. In competitions" 'prints text in display picture box
        picdisplay.Print "pros use shortboards, but also vary their board depending" 'prints text in display picture box
        picdisplay.Print "on the style of the wave." 'prints text in display picture box
        picdisplay.Print 'prints a blank line in the display picture box
        picdisplay.Print "****Click the buttons for information on these boards!****" 'prints text in display picture box and informs user of the buttons below
    End If 'ends the If stament
End Sub

Private Sub Cmdclose_Click()
    FrmBoards.Hide 'hides the boards form
End Sub

Private Sub CmdFish_Click()
    picfish.Visible = True 'shows the fishboard picture box
    picshort.Visible = False 'hides the shortboard picture box
    picsoft.Visible = False 'hides the softboard picture box
    picguns.Visible = False 'hides the guns picture box
    piclong.Visible = False 'hides the longboard picture box
    lblfish.Visible = True 'shows the fishboard label
    lblshort.Visible = False 'hides the shortboard label
    lblsoft.Visible = False 'hides the softboard label
    lblguns.Visible = False 'hides the guns label
    lbllong.Visible = False 'hides the longboard label
End Sub

Private Sub cmdguns_Click()
    picguns.Visible = True 'shows the guns picture box
    piclong.Visible = False 'hides the longboard picture box
    picsoft.Visible = False 'hides the softboard picture box
    picshort.Visible = False 'hides the shortboard picture box
    picfish.Visible = False 'hides the fishboard picture box
    lblguns.Visible = True 'shows the guns label
    lblfish.Visible = False 'hides the fishboard label
    lblshort.Visible = False 'hides the shortboard label
    lblsoft.Visible = False 'hides the softboard label
    lbllong.Visible = False 'hides the longboard label
End Sub

Private Sub cmdlongboard_Click()
    piclong.Visible = True 'Shows the longboard picture box
    picsoft.Visible = False 'hides the softboard picture box
    picshort.Visible = False 'hides the shortboard picture box
    picguns.Visible = False 'hides the guns picture box
    picfish.Visible = False 'hides the fishboard picture box
    lbllong.Visible = True 'shows the longboard label
    lblguns.Visible = False 'hides the guns label
    lblfish.Visible = False 'hides the fishboard label
    lblshort.Visible = False 'hides the shortboard label
    lblsoft.Visible = False 'hides the softboard label
End Sub

Private Sub cmdshortboard_Click()
    picshort.Visible = True 'shows the shortboard picture box
    picsoft.Visible = False 'hides the softboard picture box
    picguns.Visible = False 'hides the guns picture box
    piclong.Visible = False 'hides the longboard picture box
    picfish.Visible = False 'hides the fishboard picture box
    lblshort.Visible = True 'shows the shortboard label
    lbllong.Visible = False 'hides the longboard label
    lblguns.Visible = False 'hides the guns label
    lblfish.Visible = False 'hides the fishboard label
    lblsoft.Visible = False 'hides the softboard label
End Sub

Private Sub Cmdsoftboards_Click()
    picsoft.Visible = True 'shows the softboard picture box
    picshort.Visible = False 'hides the shortboard picture box
    picguns.Visible = False 'hides the guns picture box
    piclong.Visible = False 'hides the longboard picture box
    picfish.Visible = False 'hides the fishboard picture box
    lblsoft.Visible = True 'shows the softboard label
    lblshort.Visible = False 'hides the shortboard label
    lbllong.Visible = False 'hides the longboard label
    lblguns.Visible = False 'hides the guns label
    lblfish.Visible = False 'hides the fishboard label
End Sub

