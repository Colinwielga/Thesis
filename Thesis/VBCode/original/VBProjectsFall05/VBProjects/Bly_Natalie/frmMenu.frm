VERSION 5.00
Begin VB.Form frmMenu 
   BackColor       =   &H0000C000&
   Caption         =   "Menu"
   ClientHeight    =   4545
   ClientLeft      =   3540
   ClientTop       =   3375
   ClientWidth     =   8175
   LinkTopic       =   "Form1"
   ScaleHeight     =   4545
   ScaleWidth      =   8175
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FF8080&
      Height          =   1575
      Left            =   5760
      Picture         =   "frmMenu.frx":0000
      ScaleHeight     =   1515
      ScaleWidth      =   2115
      TabIndex        =   7
      Top             =   2760
      Width           =   2175
   End
   Begin VB.PictureBox Picture2 
      Height          =   1575
      Left            =   240
      Picture         =   "frmMenu.frx":0A9B
      ScaleHeight     =   1515
      ScaleWidth      =   2115
      TabIndex        =   6
      Top             =   2760
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      Height          =   1575
      Left            =   3000
      Picture         =   "frmMenu.frx":1D7B
      ScaleHeight     =   1515
      ScaleWidth      =   2115
      TabIndex        =   5
      Top             =   1800
      Width           =   2175
   End
   Begin VB.CommandButton cmdAbout 
      BackColor       =   &H00FF80FF&
      Caption         =   "About this Program"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CommandButton cmdGroceryStore 
      BackColor       =   &H00FF80FF&
      Caption         =   "Grocery Store"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   6000
      MaskColor       =   &H00FF80FF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton cmdMainfeature 
      BackColor       =   &H00FF80FF&
      Caption         =   "Estimate how much ""extra"" money you have this month using your most recent paycheck!"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   360
      MaskColor       =   &H00008000&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      UseMaskColor    =   -1  'True
      Width           =   1815
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FF80FF&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label lblNatalie 
      BackColor       =   &H0000C000&
      Caption         =   "by: Natalie Bly"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   8
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Welcome to the Money             Manager 2005:        Can I afford to eat this month? "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   2520
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Money Manager 2005 (ProjectNataliesMoneyPlanner)
'frmMenu (frmMenu.frm)
'by Natalie Bly
'10/29/05
'This is the start-up form for my project.  It serves as a Menu from which
'to access the features of the program.  With this project, I attempt
'to make a process that I go through each month a little easier.
'Every time I get paid, I sit down with my checkbook balance,
'my previous tuition bill, and my grocery reciepts to try to decide
'how much money from my paycheck I can spend on groceries or fun things
'and still be able to cover next semester's tuition.  I realize, however, that not
'every student faces the same financial obligations, so I tried to make the program
'flexible enough to allow for other situations.  The surplus calculator is not precise,
'but makes a projection that gives the user a place to start.


Option Explicit         'helps to debug the code
Private Sub cmdAbout_Click()
    frmMenu.Hide         'takes user to the About screen
    frmAbout.Show
End Sub

Private Sub cmdExit_Click()
    End                  'ends the program
End Sub

Private Sub cmdMainfeature_Click()
    frmMenu.Hide        'takes user back to the Menu screen
    frmBegin.Show
    cmdMainfeature.Enabled = False  'prevents user from entering the main feature screen again (to prevent strange things happening to the variables)
End Sub

Private Sub cmdGroceryStore_Click()
    Open App.Path & "\Groceries.txt" For Input As #1    'opens file in order to retrieve data
    Do Until EOF(1)                 'reads the data from the file into three arrays-an item array, a price array, and a category array
        I = I + 1
        Input #1, Item(I), Price(I), Category(I)
    Loop
    Close #1                        'closes the data file, now that the information has been stored in arrays
    frmMenu.Hide                    'takes user to the grocery store feature
    frmGroceryStore.Show
End Sub
