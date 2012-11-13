VERSION 5.00
Begin VB.Form frmStart 
   BackColor       =   &H00008000&
   Caption         =   "Start Page"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Cmddirectory 
      Caption         =   "Directory"
      Height          =   855
      Left            =   6000
      TabIndex        =   12
      Top             =   1560
      Width           =   2895
   End
   Begin VB.PictureBox Picture4 
      Height          =   1815
      Left            =   12480
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   1755
      ScaleWidth      =   2235
      TabIndex        =   11
      Top             =   4560
      Width           =   2295
   End
   Begin VB.PictureBox Picture7 
      Height          =   1575
      Left            =   12360
      Picture         =   "Form1.frx":CB4E
      ScaleHeight     =   1515
      ScaleWidth      =   1995
      TabIndex        =   10
      Top             =   7080
      Width           =   2055
   End
   Begin VB.PictureBox Picture11 
      Height          =   1455
      Left            =   360
      Picture         =   "Form1.frx":17154
      ScaleHeight     =   1395
      ScaleWidth      =   1755
      TabIndex        =   9
      Top             =   600
      Width           =   1815
   End
   Begin VB.PictureBox Picture10 
      Height          =   2055
      Left            =   9240
      Picture         =   "Form1.frx":48396
      ScaleHeight     =   1995
      ScaleWidth      =   2235
      TabIndex        =   8
      Top             =   7080
      Width           =   2295
   End
   Begin VB.PictureBox Picture9 
      Height          =   1935
      Left            =   3360
      Picture         =   "Form1.frx":56B24
      ScaleHeight     =   1875
      ScaleWidth      =   1875
      TabIndex        =   7
      Top             =   7200
      Width           =   1935
   End
   Begin VB.PictureBox Picture8 
      Height          =   1815
      Left            =   6240
      Picture         =   "Form1.frx":623E6
      ScaleHeight     =   1755
      ScaleWidth      =   1995
      TabIndex        =   6
      Top             =   7200
      Width           =   2055
   End
   Begin VB.PictureBox Picture6 
      Height          =   1575
      Left            =   12600
      Picture         =   "Form1.frx":6DED8
      ScaleHeight     =   1515
      ScaleWidth      =   1995
      TabIndex        =   5
      Top             =   2280
      Width           =   2055
   End
   Begin VB.PictureBox Picture5 
      Height          =   1935
      Left            =   240
      Picture         =   "Form1.frx":7800A
      ScaleHeight     =   1875
      ScaleWidth      =   1875
      TabIndex        =   4
      Top             =   2280
      Width           =   1935
   End
   Begin VB.PictureBox Picture3 
      Height          =   1575
      Left            =   12600
      Picture         =   "Form1.frx":843D0
      ScaleHeight     =   1515
      ScaleWidth      =   1635
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
   Begin VB.PictureBox Picture2 
      Height          =   2295
      Left            =   240
      Picture         =   "Form1.frx":8BCBA
      ScaleHeight     =   2235
      ScaleWidth      =   2235
      TabIndex        =   2
      Top             =   4560
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      Height          =   1815
      Left            =   840
      Picture         =   "Form1.frx":9C5D4
      ScaleHeight     =   1755
      ScaleWidth      =   1395
      TabIndex        =   1
      Top             =   7320
      Width           =   1455
   End
   Begin VB.Label Label10 
      BackColor       =   &H00008000&
      Caption         =   "An inforamtional center is also available to see what Cross Country is all about!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   21
      Top             =   3720
      Width           =   6975
   End
   Begin VB.Label Label5 
      BackColor       =   &H00008000&
      Caption         =   "Another feature that is offered is that you are able to view the ALL of the MIAC schools CC websites!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   20
      Top             =   6240
      Width           =   8775
   End
   Begin VB.Label Label4 
      BackColor       =   &H00008000&
      Caption         =   "An additional feature that this program is that there is a mile/kilometer converter along with a Pace calculator!! "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   19
      Top             =   4560
      Width           =   9615
   End
   Begin VB.Label Label3 
      BackColor       =   &H00008000&
      Caption         =   "You are able to see individual results along with team results."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   18
      Top             =   5400
      Width           =   5535
   End
   Begin VB.Label Label2 
      BackColor       =   &H00008000&
      Caption         =   "This project gives a different perspective on viewing the results of the 2008 MIAC Men's Conference Meet.  "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   17
      Top             =   3000
      Width           =   9375
   End
   Begin VB.Label Label8 
      BackColor       =   &H00008000&
      Caption         =   "By: Tyler Trettel and Josh Gunderson"
      Height          =   255
      Left            =   2520
      TabIndex        =   16
      Top             =   10080
      Width           =   2895
   End
   Begin VB.Label Label6 
      BackColor       =   &H00008000&
      Caption         =   "2008 MIAC Cross Country Project "
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   15
      Top             =   10080
      Width           =   2535
   End
   Begin VB.Label Label7 
      BackColor       =   &H00008000&
      Caption         =   "About the Program/ MIAC Logos"
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   10320
      Width           =   2535
   End
   Begin VB.Label Label9 
      BackColor       =   &H00008000&
      Caption         =   "November 5, 2008"
      Height          =   255
      Left            =   2520
      TabIndex        =   13
      Top             =   10320
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackColor       =   &H00008000&
      Caption         =   "2008 Men MIAC Cross Country"
      BeginProperty Font 
         Name            =   "Rockwell Extra Bold"
         Size            =   27.75
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2520
      TabIndex        =   0
      Top             =   240
      Width           =   9855
   End
End
Attribute VB_Name = "frmstart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project Name: MIAC CC Project
'Form Name: frmStart
'Authors: Josh Gunderson & Tyler Trettel
'Date: 5 November 2008
'Objective: The purpose of this form is for the user to view each of teh schools logo along with it will lead them to the directory

Private Sub cmddirectory_Click()
frmstart.Hide
frmdirectory.Show

End Sub

