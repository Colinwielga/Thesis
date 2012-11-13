VERSION 5.00
Begin VB.Form Victoria 
   BackColor       =   &H00FF8080&
   Caption         =   "Victoria"
   ClientHeight    =   13725
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19080
   LinkTopic       =   "Victoria"
   ScaleHeight     =   13725
   ScaleWidth      =   19080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12120
      TabIndex        =   15
      Top             =   11040
      Width           =   1695
   End
   Begin VB.CommandButton cmdreturn 
      Caption         =   "Return to Map of London"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   9960
      TabIndex        =   14
      Top             =   10560
      Width           =   1815
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   720
      TabIndex        =   13
      Text            =   "Victoria Station"
      Top             =   6720
      Width           =   1815
   End
   Begin VB.PictureBox picStation 
      Height          =   3015
      Left            =   600
      Picture         =   "Victoria.frx":0000
      ScaleHeight     =   2955
      ScaleWidth      =   13995
      TabIndex        =   12
      Top             =   7200
      Width           =   14055
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8040
      TabIndex        =   11
      Text            =   "Westminister Cathedral"
      Top             =   120
      Width           =   3375
   End
   Begin VB.PictureBox picCathedral 
      Height          =   6615
      Left            =   7080
      Picture         =   "Victoria.frx":9DF9
      ScaleHeight     =   6555
      ScaleWidth      =   5355
      TabIndex        =   10
      Top             =   480
      Width           =   5415
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   360
      TabIndex        =   9
      Text            =   "Westminister Abbey"
      Top             =   2640
      Width           =   2175
   End
   Begin VB.PictureBox picabbey 
      Height          =   3375
      Left            =   240
      Picture         =   "Victoria.frx":11D3D
      ScaleHeight     =   3315
      ScaleWidth      =   4755
      TabIndex        =   8
      Top             =   3000
      Width           =   4815
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   840
      TabIndex        =   6
      Text            =   "Victoria Station"
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   720
      TabIndex        =   5
      Text            =   "Westminitster Cathedral"
      Top             =   1440
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   720
      TabIndex        =   4
      Text            =   "Westminister Abbey"
      Top             =   960
      Width           =   2175
   End
   Begin VB.OptionButton optstation 
      Caption         =   "Option3"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1920
      Width           =   255
   End
   Begin VB.OptionButton optcathedral 
      Caption         =   "Option2"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1440
      Width           =   255
   End
   Begin VB.OptionButton optabbey 
      Caption         =   "Option1"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "Created by Chelsey Johnson"
      Height          =   255
      Left            =   360
      TabIndex        =   16
      Top             =   11760
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF80&
      Caption         =   "Choose one of the sites to learn more about the famous site."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   600
      Width           =   5775
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "These pictures are all of famous places in London, found in the Victoria district."
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7215
   End
End
Attribute VB_Name = "Victoria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Discovering London(Project1.vbp)
'Form Name: Victoria (Victoria.frm)
'Author: Chelsey Johnson
'Date Written: March 14, 2004
'Purpose of Form: This form is to able to let the user choose which site they would like to learn about by clicking on an option button.
                    'The purpose is for them to have the choice of learning about The Westminister Abbey, The Westminister
                    'Cathedral, The Victoria Station, or all of them
'Option Explicit is a command to force
'the user to explicitly declare all variables
'before they can be used.
Option Explicit

Private Sub cmdquit_Click()
End
End Sub

Private Sub cmdreturn_Click()
'User returns of the Map of London page to choose a new variable
Victoria.Hide
MapLondon.Show
End Sub

Private Sub optabbey_Click()
'By choosing this option the user learns about The Westminister Abbey
MsgBox "An architectural masterpiece of the thirteenth to sixteenth centuries,Westminster Abbey also presents a unique pageant of British history -the Confessor’s Shrine, the tombs of Kings and Queens,and countless memorials to the famous and the great.It has been the setting for every Coronation since 1066 and for numerous other Royal occasions. Today it is still a church dedicated to regular worship and to the celebration of great events in the life of the nation. Neither a cathedral nor a parish church, Westminster Abbey is a “royal peculiar” under the jurisdiction of a Dean and Chapter, subject only to the Sovereign.", , "Westminister Abbey"
End Sub

Private Sub optcathedral_Click()
'By choosing this option the user learns about The Westminister Cathedral
MsgBox "The Cathedral Church of Westminster, which is dedicated to the Most Precious Blood of Our Lord Jesus Christ, was designed in the Early Christian Byzantine style by the Victorian architect John Francis Bentley. The foundation stone was laid in 1895 and the fabric of the building was completed eight years later.", , "Westminister Cathedral"
End Sub

Private Sub optstation_Click()
'By choosing this option the user learns about Victoria Station
MsgBox "Trains leave here for the south Coast - Brighton, Dover, eastbourne.   The station has an impressive glazed ironwork roof.", , "Victoria Station"
End Sub
