VERSION 5.00
Begin VB.Form apparel 
   BackColor       =   &H0000C000&
   Caption         =   "Apparel"
   ClientHeight    =   7005
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11175
   FillColor       =   &H0000C000&
   ForeColor       =   &H0000C000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7005
   ScaleWidth      =   11175
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton advance 
      Caption         =   "NEXT"
      Height          =   615
      Left            =   6720
      TabIndex        =   9
      Top             =   1800
      Width           =   1095
   End
   Begin VB.PictureBox picResults 
      Height          =   1455
      Left            =   8040
      ScaleHeight     =   1395
      ScaleWidth      =   2835
      TabIndex        =   8
      Top             =   1080
      Width           =   2895
   End
   Begin VB.CommandButton selected 
      BackColor       =   &H00800080&
      Caption         =   "Apparel Selected"
      Height          =   615
      Left            =   6720
      MaskColor       =   &H0000C000&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1080
      Width           =   1095
   End
   Begin VB.OptionButton optlongsleevetshirt 
      BackColor       =   &H0000C000&
      Caption         =   "LONG SLEEVE T-SHIRT"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   2880
      Width           =   2295
   End
   Begin VB.OptionButton opthoodedsweatshirt 
      BackColor       =   &H0000C000&
      Caption         =   "HOODED SWEATSHIRT"
      Height          =   495
      Left            =   6360
      TabIndex        =   4
      Top             =   3720
      Width           =   1455
   End
   Begin VB.OptionButton optcrewsweatshirt 
      BackColor       =   &H0000C000&
      Caption         =   "CREW SWEATSHIRT"
      Height          =   255
      Left            =   8400
      TabIndex        =   3
      Top             =   3840
      Width           =   2055
   End
   Begin VB.OptionButton optshorts 
      BackColor       =   &H0000C000&
      Caption         =   "SHORTS"
      Height          =   255
      Left            =   3480
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.OptionButton opttshirt 
      BackColor       =   &H0000C000&
      Caption         =   "T-SHIRT"
      Height          =   195
      Left            =   3960
      TabIndex        =   1
      Top             =   3600
      Width           =   975
   End
   Begin VB.OptionButton optsweatpants 
      BackColor       =   &H0000C000&
      Caption         =   "SWEATPANTS"
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   3240
      Top             =   1680
      Width           =   15
   End
   Begin VB.Label select 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "PLEASE SELECT APPAREL TYPE"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   5160
      TabIndex        =   6
      Top             =   960
      Width           =   1695
   End
   Begin VB.Image crewsweatshirt 
      Height          =   1830
      Left            =   8760
      Picture         =   "apparel.frx":0000
      Top             =   4320
      Width           =   1500
   End
   Begin VB.Image shorts 
      Height          =   2070
      Left            =   3000
      Picture         =   "apparel.frx":597C
      Top             =   600
      Width           =   2025
   End
   Begin VB.Image hoodedsweatshirt 
      Height          =   2235
      Left            =   6360
      Picture         =   "apparel.frx":6382
      Top             =   4320
      Width           =   1500
   End
   Begin VB.Image longsleevetshirt 
      Height          =   3750
      Left            =   480
      Picture         =   "apparel.frx":CA6D
      Top             =   3240
      Width           =   2430
   End
   Begin VB.Image tshirt 
      Height          =   2235
      Left            =   3720
      Picture         =   "apparel.frx":F592
      Top             =   3960
      Width           =   1500
   End
   Begin VB.Image sweatpants 
      Height          =   2235
      Left            =   720
      Picture         =   "apparel.frx":14FF0
      Top             =   480
      Width           =   1500
   End
End
Attribute VB_Name = "apparel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : Screen Printing(Main1.vpb)
'Form Name : Apparel (apparel.frm)
'Author: Elizabeth Beesch
'Date Written: October 27, 2003
'Purpose of Form: To have the user select a type of apparel, calculate the cost
        ' of the apparel and then print it again on the final form
'Option Explicit is a command to force
'the user to explicitly declare all variables
'before they can be used.
Option Explicit
Private Sub Form_Load() 'when the form loads the following will happen
advance.Enabled = False ' dosen't allow use of this button
End Sub 'end of operations for this sub
Private Sub optcrewsweatshirt_Click() 'when this button is clicked on the following will happen
A = 16 ' the variable will equal this number
End Sub ' end of operations for this sub
Private Sub opthoodedsweatshirt_Click() 'when this button is clicked on the following will happen
A = 18 ' the variable will equal this number
End Sub ' end of operations for this sub
Private Sub optlongsleevetshirt_Click() 'when this button is clicked on the following will happen
A = 8 ' the variable will equal this number
End Sub ' end of operations for this sub
Private Sub optshorts_Click() 'when this button is clicked on the following will happen
A = 10 ' the variable will equal this number
End Sub ' end of operations for this sub
Private Sub optsweatpants_Click() 'when this button is clicked on the following will happen
A = 12 ' the variable will equal this number
End Sub ' end of operations for this sub
Private Sub opttshirt_Click() 'when this button is clicked on the following will happen
A = 6 ' the variable will equal this number
End Sub ' end of operations for this sub
Private Sub selected_Click() 'when this button is clicked on the following will happen
apparelcost = A * numberofItems ' calculates the apparel cost
picResults.Print "Total Cost for Apparel:"; FormatCurrency(apparelcost, 2) ' prints the apparel cost
selected.Enabled = False 'dosen't allow use of this button
advance.Enabled = True ' allows use of this button
End Sub ' end of operations for this sub
Private Sub advance_Click() 'when this button is clicked on the following will happen
apparel.Hide 'hides current form
jerseyname.Show 'shows next form in sequence
End Sub ' end of operations for this sub
