VERSION 5.00
Begin VB.Form frmhats 
   BackColor       =   &H00FF8080&
   Caption         =   "Hats"
   ClientHeight    =   7800
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11040
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7800
   ScaleWidth      =   11040
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdfinalize 
      BackColor       =   &H008080FF&
      Caption         =   "Finalize Your Order"
      Height          =   735
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6000
      Width           =   2295
   End
   Begin VB.CommandButton cmdclearhats 
      BackColor       =   &H008080FF&
      Caption         =   "Clear Calcutated Price of Hats"
      Height          =   735
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6000
      Width           =   2295
   End
   Begin VB.CommandButton cmdcalchats 
      BackColor       =   &H008080FF&
      Caption         =   "Calculate Price of Hats"
      Height          =   735
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4920
      Width           =   2295
   End
   Begin VB.TextBox txthats 
      Height          =   495
      Left            =   3960
      TabIndex        =   13
      Top             =   4080
      Width           =   2175
   End
   Begin VB.PictureBox picresultshat 
      BackColor       =   &H8000000E&
      Height          =   3015
      Left            =   6840
      ScaleHeight     =   2955
      ScaleWidth      =   3675
      TabIndex        =   8
      Top             =   4080
      Width           =   3735
   End
   Begin VB.OptionButton opt2colorhat 
      BackColor       =   &H00FF8080&
      Caption         =   "Option4"
      Height          =   255
      Left            =   9360
      TabIndex        =   7
      Top             =   240
      Width           =   255
   End
   Begin VB.OptionButton optnotfittedhat 
      BackColor       =   &H00FF8080&
      Caption         =   "Option3"
      Height          =   255
      Left            =   6600
      TabIndex        =   6
      Top             =   240
      Width           =   255
   End
   Begin VB.OptionButton optretrohat 
      BackColor       =   &H00FF8080&
      Caption         =   "Option2"
      Height          =   255
      Left            =   3960
      TabIndex        =   5
      Top             =   240
      Width           =   255
   End
   Begin VB.OptionButton optlogohat 
      BackColor       =   &H00FF8080&
      Caption         =   "Option1"
      Height          =   255
      Left            =   1320
      TabIndex        =   4
      Top             =   240
      Width           =   255
   End
   Begin VB.PictureBox pic2colorhat 
      Height          =   1815
      Left            =   8400
      Picture         =   "hats.frx":0000
      ScaleHeight     =   1755
      ScaleWidth      =   2115
      TabIndex        =   3
      Top             =   600
      Width           =   2175
   End
   Begin VB.PictureBox picnotfittedhat 
      Height          =   1815
      Left            =   5640
      Picture         =   "hats.frx":DAE2
      ScaleHeight     =   1755
      ScaleWidth      =   2115
      TabIndex        =   2
      Top             =   600
      Width           =   2175
   End
   Begin VB.PictureBox picretrohat 
      Height          =   1815
      Left            =   3000
      Picture         =   "hats.frx":1B5C4
      ScaleHeight     =   1755
      ScaleWidth      =   2115
      TabIndex        =   1
      Top             =   600
      Width           =   2175
   End
   Begin VB.PictureBox piclogohat 
      Height          =   1815
      Left            =   360
      Picture         =   "hats.frx":290A6
      ScaleHeight     =   1755
      ScaleWidth      =   2115
      TabIndex        =   0
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "23.95"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   8280
      TabIndex        =   21
      Top             =   3120
      Width           =   2415
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "$14.99"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   5520
      TabIndex        =   20
      Top             =   3120
      Width           =   2415
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "$21.95"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   2880
      TabIndex        =   19
      Top             =   3240
      Width           =   2415
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "$18.95"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   495
      Left            =   240
      TabIndex        =   18
      Top             =   3120
      Width           =   2415
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "How Many Hats??"
      BeginProperty Font 
         Name            =   "Mathematica6"
         Size            =   15.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   615
      Left            =   1920
      TabIndex        =   14
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "Two Colors"
      BeginProperty Font 
         Name            =   "Maiandra GD"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   8400
      TabIndex        =   12
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "Adjustable"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   5640
      TabIndex        =   11
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "Team Nickname"
      BeginProperty Font 
         Name            =   "Rockwell Extra Bold"
         Size            =   12
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Left            =   3000
      TabIndex        =   10
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "Team Logo"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   9
      Top             =   2640
      Width           =   2175
   End
End
Attribute VB_Name = "frmhats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : BaseballUniforms (BaseballUniforms.vbp)
'Form Name : frmhats (hats.frm)
'Author: Kyle Kaczmarek
'Date Written: March 15, 2004
'Purpose of Form:'

Private Sub cmdcalchats_Click()


hats = txthats.Text 'input how may hats you want


If optlogohat = True Then 'chosen logo hat option
        hatsprice = 18.95 'logo hat price
    ElseIf optretrohat = True Then 'chosen retro hat option
        hatsprice = 21.95 'retro het price
    ElseIf optnotfittedhat = True Then 'chosen not fitted hat
        hatsprice = 14.99 'not fitted hat price
    ElseIf opt2colorhat = True Then 'chosen 2 color hat
        hatsprice = 23.95 '2 c olor hat price
End If

hatstotal = hats * hatsprice 'multiplies the number of hats ordered by the price
    
picresultshat.Cls 'clears the picture box
picresultshat.Print "Number of Hats", "Cost" 'prints the titles
picresultshat.Print "***********************", "*****" 'prints the stars
picresultshat.Print Tab(8); hats, , FormatCurrency(hatstotal, 2) 'prints the total
End Sub

Private Sub cmdclearhats_Click()

picresultshat.Cls 'clears the picture box

End Sub

Private Sub cmdfinalize_Click()
frmcleats.Hide 'closes the cleats form
frmjerseys.Hide 'closes the jerseys form
frmhats.Hide 'closes the hat form
frmorder.Hide 'closes the order form
frmpants.Hide 'closes the pants form
frmfinal.Show 'shows the final form
End Sub
